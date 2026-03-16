// Copyright 2016 - 2025 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package excelize

import (
	"context"
	"database/sql"
	"errors"
	"fmt"
	"log"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"time"

	"github.com/jackc/pgx/v5"
	"github.com/jackc/pgx/v5/stdlib"
)

var pgIdentRegexp = regexp.MustCompile(`^[A-Za-z_][A-Za-z0-9_]*$`)

// PGSyncOptions defines options for syncing workbook data into PostgreSQL.
type PGSyncOptions struct {
	// WorkbookID is the logical ID of the workbook in PostgreSQL.
	// If empty, base name from File.Path is used, and falls back to "workbook".
	WorkbookID string
	// Schema is the target schema. Default is "public".
	Schema string
	// TablePrefix is prefixed to mirror tables. Default is "excelize".
	TablePrefix string
	// BatchSize controls INSERT batch size for cell upsert. Default is 1000.
	BatchSize int
	// RawCellValue controls whether GetRows uses raw values during full sync.
	RawCellValue bool
	// AutoWarmup preloads supported PostgreSQL lookup caches during EnablePostgresMirror.
	AutoWarmup bool
	// WarmupTimeout bounds the AutoWarmup phase. Zero means no timeout.
	WarmupTimeout time.Duration
	// WarmupSheets limits AutoWarmup to specific worksheets. Empty means all sheets.
	WarmupSheets []string
	// AsyncWarmup runs AutoWarmup in the background so EnablePostgresMirror returns immediately.
	AsyncWarmup bool
}

// PGCellSnapshot is a single mirrored cell from PostgreSQL.
type PGCellSnapshot struct {
	WorkbookID string
	Sheet      string
	Cell       string
	Row        int
	Col        int
	Value      string
	ValueType  string
	Formula    string
	UpdatedAt  time.Time
}

type pgMirrorPendingCell struct {
	HasValue  bool
	Value     string
	ValueType string
}

// PGWarmupStatus describes the latest PostgreSQL lookup warmup state.
type PGWarmupStatus struct {
	Enabled     bool
	InProgress  bool
	Async       bool
	StartedAt   time.Time
	CompletedAt time.Time
	Sheets      []string
	Err         error
}

type pgMirrorTables struct {
	workbook string
	sheet    string
	cell     string
	lookup   string
}

type pgMirrorCell struct {
	Cell      string
	Row       int
	Col       int
	Value     string
	ValueType string
	Formula   string
}

type pgMirrorSheetMeta struct {
	RowCount int
	ColCount int
}

type pgMirrorCellRef struct {
	Sheet string
	Cell  string
}

type pgMirrorPendingCellRef struct {
	Sheet     string
	Cell      string
	HasValue  bool
	Value     string
	ValueType string
}

type pgMirrorExecutor interface {
	ExecContext(context.Context, string, ...interface{}) (sql.Result, error)
}

var pgMirrorCopyFromTestHook func(sheet string, rowCount int)
var pgMirrorWarmupTestHook func()

// SyncWorkbookToPostgres mirrors all worksheet cells from this workbook into PostgreSQL.
//
// The mirror schema includes three tables:
//  1. <prefix>_workbook_mirror
//  2. <prefix>_sheet_mirror
//  3. <prefix>_cell_mirror
//
// Existing mirrored cells for each sheet are replaced during sync.
func (f *File) SyncWorkbookToPostgres(ctx context.Context, db *sql.DB, opts ...PGSyncOptions) error {
	if db == nil {
		return errors.New("postgres mirror requires non-nil *sql.DB")
	}
	cfg, err := f.getPGSyncOptions(opts...)
	if err != nil {
		return err
	}
	if ctx == nil {
		ctx = context.Background()
	}
	if sheetMeta, handled, err := f.trySyncWorkbookToPostgresViaPGXCopy(ctx, db, cfg); handled {
		if err != nil {
			return err
		}
		f.replacePGMirrorSheetMeta(sheetMeta)
		return nil
	}
	tx, err := db.BeginTx(ctx, nil)
	if err != nil {
		return err
	}
	rollback := true
	defer func() {
		if rollback {
			_ = tx.Rollback()
		}
	}()

	tables := getPGMirrorTables(cfg.TablePrefix)
	sheets := f.GetSheetList()
	sheetMeta := make(map[string]pgMirrorSheetMeta, len(sheets))
	if err = ensurePGMirrorTables(ctx, tx, cfg.Schema, tables); err != nil {
		return err
	}
	if err = upsertPGWorkbookMirror(ctx, tx, cfg.Schema, tables, cfg.WorkbookID, f.Path); err != nil {
		return err
	}
	for _, sheet := range sheets {
		meta, syncErr := f.syncSheetToPostgresTx(ctx, tx, cfg, tables, sheet)
		if syncErr != nil {
			err = syncErr
			return err
		}
		sheetMeta[sheet] = meta
	}
	if err = deletePGRemovedSheets(ctx, tx, cfg.Schema, tables, cfg.WorkbookID, sheets); err != nil {
		return err
	}
	if err = tx.Commit(); err != nil {
		return err
	}
	rollback = false
	f.replacePGMirrorSheetMeta(sheetMeta)
	return nil
}

// EnablePostgresMirror enables automatic PostgreSQL mirror sync for cell writes.
//
// When bestEffort is false, direct SetCellValue / SetCellFormula mutations return
// mirror errors after the workbook mutation is applied. Recalculation-triggered
// mirror writes can't propagate errors, so they are logged instead.
func (f *File) EnablePostgresMirror(db *sql.DB, bestEffort bool, opts ...PGSyncOptions) error {
	if db == nil {
		return errors.New("postgres mirror requires non-nil *sql.DB")
	}
	cfg, err := f.getPGSyncOptions(opts...)
	if err != nil {
		return err
	}
	f.mu.Lock()
	prevEnabled := f.pgMirrorEnabled
	prevBestEffort := f.pgMirrorBestEffort
	prevDB := f.pgMirrorDB
	prevOpts := clonePGSyncOptions(f.pgMirrorOpts)
	f.pgMirrorEnabled = true
	f.pgMirrorBestEffort = bestEffort
	f.pgMirrorDB = db
	f.pgMirrorOpts = clonePGSyncOptions(cfg)
	if !cfg.AutoWarmup {
		f.resetPGWarmupLocked(true)
		f.mu.Unlock()
		return nil
	}
	warmupID, warmupCtx := f.startPGWarmupLocked(cfg)
	f.mu.Unlock()

	if cfg.AsyncWarmup {
		go f.runPostgresWarmup(warmupID, warmupCtx, cfg)
		return nil
	}

	err = f.preloadPostgresWarmup(warmupCtx, cfg)
	f.finishPGWarmup(warmupID, err)
	if err == nil {
		return nil
	}

	f.mu.Lock()
	if f.pgWarmupSeq == warmupID {
		f.pgMirrorEnabled = prevEnabled
		f.pgMirrorBestEffort = prevBestEffort
		f.pgMirrorDB = prevDB
		f.pgMirrorOpts = prevOpts
		status := f.pgWarmupStatus
		status.Enabled = prevEnabled && prevDB != nil && prevOpts != nil
		f.pgWarmupStatus = clonePGWarmupStatus(status)
	}
	f.mu.Unlock()
	return err
}

// DisablePostgresMirror disables automatic PostgreSQL mirror sync.
func (f *File) DisablePostgresMirror() {
	f.mu.Lock()
	defer f.mu.Unlock()
	f.resetPGWarmupLocked(false)
	f.pgMirrorEnabled = false
	f.pgMirrorBestEffort = false
	f.pgMirrorDB = nil
	f.pgMirrorOpts = nil
}

// GetPostgresWarmupStatus returns the latest PostgreSQL lookup warmup status.
func (f *File) GetPostgresWarmupStatus() PGWarmupStatus {
	f.mu.Lock()
	defer f.mu.Unlock()
	return clonePGWarmupStatus(f.pgWarmupStatus)
}

// WaitForPostgresWarmup waits for the current PostgreSQL lookup warmup to finish.
func (f *File) WaitForPostgresWarmup(ctx context.Context) error {
	if ctx == nil {
		ctx = context.Background()
	}
	for {
		f.mu.Lock()
		done := f.pgWarmupDone
		status := clonePGWarmupStatus(f.pgWarmupStatus)
		f.mu.Unlock()

		if done == nil || !status.InProgress {
			return status.Err
		}
		select {
		case <-done:
		case <-ctx.Done():
			return ctx.Err()
		}
	}
}

// SyncSheetToPostgres mirrors a single worksheet into PostgreSQL.
func (f *File) SyncSheetToPostgres(ctx context.Context, db *sql.DB, sheet string, opts ...PGSyncOptions) error {
	if db == nil {
		return errors.New("postgres mirror requires non-nil *sql.DB")
	}
	if err := checkSheetName(sheet); err != nil {
		return err
	}
	cfg, err := f.getPGSyncOptions(opts...)
	if err != nil {
		return err
	}
	if ctx == nil {
		ctx = context.Background()
	}
	if meta, handled, err := f.trySyncSheetToPostgresViaPGXCopy(ctx, db, sheet, cfg); handled {
		if err != nil {
			return err
		}
		f.mergePGMirrorSheetMeta(map[string]pgMirrorSheetMeta{
			sheet: meta,
		})
		return nil
	}
	tx, err := db.BeginTx(ctx, nil)
	if err != nil {
		return err
	}
	rollback := true
	defer func() {
		if rollback {
			_ = tx.Rollback()
		}
	}()

	tables := getPGMirrorTables(cfg.TablePrefix)
	if err = ensurePGMirrorTables(ctx, tx, cfg.Schema, tables); err != nil {
		return err
	}
	if err = upsertPGWorkbookMirror(ctx, tx, cfg.Schema, tables, cfg.WorkbookID, f.Path); err != nil {
		return err
	}
	meta, syncErr := f.syncSheetToPostgresTx(ctx, tx, cfg, tables, sheet)
	if syncErr != nil {
		err = syncErr
		return err
	}
	if err = tx.Commit(); err != nil {
		return err
	}
	rollback = false
	f.mergePGMirrorSheetMeta(map[string]pgMirrorSheetMeta{
		sheet: meta,
	})
	return nil
}

// SyncCellToPostgres mirrors one cell into PostgreSQL with upsert semantics.
func (f *File) SyncCellToPostgres(ctx context.Context, db *sql.DB, sheet, cell string, opts ...PGSyncOptions) error {
	if db == nil {
		return errors.New("postgres mirror requires non-nil *sql.DB")
	}
	if err := checkSheetName(sheet); err != nil {
		return err
	}
	col, row, err := CellNameToCoordinates(cell)
	if err != nil {
		return err
	}
	cfg, err := f.getPGSyncOptions(opts...)
	if err != nil {
		return err
	}
	if ctx == nil {
		ctx = context.Background()
	}
	tx, err := db.BeginTx(ctx, nil)
	if err != nil {
		return err
	}
	rollback := true
	defer func() {
		if rollback {
			_ = tx.Rollback()
		}
	}()

	tables := getPGMirrorTables(cfg.TablePrefix)
	if err = ensurePGMirrorTables(ctx, tx, cfg.Schema, tables); err != nil {
		return err
	}
	if err = upsertPGWorkbookMirror(ctx, tx, cfg.Schema, tables, cfg.WorkbookID, f.Path); err != nil {
		return err
	}

	value, err := f.GetCellValue(sheet, cell, Options{RawCellValue: cfg.RawCellValue})
	if err != nil {
		return err
	}
	formula, err := f.GetCellFormula(sheet, cell)
	if err != nil {
		return err
	}
	cellType, err := f.GetCellType(sheet, cell)
	if err != nil {
		return err
	}
	valueType := pgMirrorValueTypeFromCell(value, cellType)

	// Keep sheet metadata monotonic for dimensions.
	if err = upsertPGSheetMirror(ctx, tx, cfg.Schema, tables, cfg.WorkbookID, sheet, row, col, -1); err != nil {
		return err
	}

	if value == "" && formula == "" {
		if err = deletePGMirroredCell(ctx, tx, cfg.Schema, tables, cfg.WorkbookID, sheet, cell); err != nil {
			return err
		}
		if err = deletePGLookupRow(ctx, tx, cfg.Schema, tables, cfg.WorkbookID, sheet, cell); err != nil {
			return err
		}
	} else {
		cellRecord := pgMirrorCell{
			Cell:      cell,
			Row:       row,
			Col:       col,
			Value:     value,
			ValueType: valueType,
			Formula:   formula,
		}
		if err = upsertPGMirroredCells(ctx, tx, cfg.Schema, tables, cfg.WorkbookID, sheet, []pgMirrorCell{cellRecord}); err != nil {
			return err
		}
		if err = upsertPGLookupRows(ctx, tx, cfg.Schema, tables, cfg.WorkbookID, sheet, []pgMirrorCell{cellRecord}); err != nil {
			return err
		}
	}

	if err = tx.Commit(); err != nil {
		return err
	}
	rollback = false
	f.updatePGMirrorSheetMeta(sheet, row, col)
	f.markPGMirrorReady(true)
	return nil
}

// LoadCellSnapshotFromPostgres loads a mirrored cell snapshot from PostgreSQL.
func (f *File) LoadCellSnapshotFromPostgres(ctx context.Context, db *sql.DB, sheet, cell string, opts ...PGSyncOptions) (PGCellSnapshot, error) {
	if db == nil {
		return PGCellSnapshot{}, errors.New("postgres mirror requires non-nil *sql.DB")
	}
	if err := checkSheetName(sheet); err != nil {
		return PGCellSnapshot{}, err
	}
	if _, _, err := CellNameToCoordinates(cell); err != nil {
		return PGCellSnapshot{}, err
	}
	cfg, err := f.getPGSyncOptions(opts...)
	if err != nil {
		return PGCellSnapshot{}, err
	}
	if ctx == nil {
		ctx = context.Background()
	}
	tables := getPGMirrorTables(cfg.TablePrefix)
	sqlStr := fmt.Sprintf(
		`SELECT workbook_id, sheet_name, cell_ref, row_num, col_num, value, value_type, formula, updated_at
FROM %s
WHERE workbook_id = $1 AND sheet_name = $2 AND cell_ref = $3`,
		qualifyPGTable(cfg.Schema, tables.cell),
	)
	var snapshot PGCellSnapshot
	if err = db.QueryRowContext(ctx, sqlStr, cfg.WorkbookID, sheet, cell).Scan(
		&snapshot.WorkbookID, &snapshot.Sheet, &snapshot.Cell,
		&snapshot.Row, &snapshot.Col, &snapshot.Value, &snapshot.ValueType, &snapshot.Formula, &snapshot.UpdatedAt,
	); err != nil {
		return PGCellSnapshot{}, err
	}
	return snapshot, nil
}

// RefreshCellFromPostgresAndRecalculate loads a mirrored cell from PostgreSQL into workbook memory
// and recalculates the worksheet by dependency-aware DAG engine.
func (f *File) RefreshCellFromPostgresAndRecalculate(ctx context.Context, db *sql.DB, sheet, cell string, opts ...PGSyncOptions) (string, error) {
	cfg, err := f.getPGSyncOptions(opts...)
	if err != nil {
		return "", err
	}
	snapshot, err := f.LoadCellSnapshotFromPostgres(ctx, db, sheet, cell, *cfg)
	if err != nil {
		return "", err
	}
	if same, err := f.postgresCellSnapshotMatchesWorkbook(sheet, cell, snapshot); err == nil && same {
		if snapshot.Formula == "" {
			return f.renderPostgresSnapshotValue(snapshot), nil
		}
		return f.GetCellValue(sheet, cell)
	}
	if err = f.applyPostgresCellSnapshot(sheet, cell, snapshot); err != nil {
		return "", err
	}
	bypassPGLookup := snapshot.Formula != ""
	if !bypassPGLookup {
		if err = f.syncCellToPostgresFastPath(ctx, db, sheet, cell, *cfg); err != nil {
			if err = f.SyncCellToPostgres(ctx, db, sheet, cell, *cfg); err != nil {
				return "", err
			}
		}
	}
	if bypassPGLookup {
		f.mu.Lock()
		f.pgLookupBypass++
		f.mu.Unlock()
		defer func() {
			f.mu.Lock()
			if f.pgLookupBypass > 0 {
				f.pgLookupBypass--
			}
			f.mu.Unlock()
		}()
	}
	f.recalcMu.Lock()
	defer f.recalcMu.Unlock()
	f.beginPGMirrorCalculationBatch()
	defer func() {
		if flushErr := f.flushPGMirrorCalculationBatch(); flushErr != nil {
			log.Printf("excelize: postgres mirror batch sync failed after RefreshCellFromPostgresAndRecalculate: %v", flushErr)
		}
	}()
	if !f.recalculateDirtyValueSubgraph(f.snapshotDirtyValueCells(), "PGRefresh") {
		if err = f.recalculateAllWithDependencyLocked(); err != nil {
			return "", err
		}
	}
	return f.GetCellValue(sheet, cell)
}

func (f *File) postgresCellSnapshotMatchesWorkbook(sheet, cell string, snapshot PGCellSnapshot) (bool, error) {
	currentFormula, err := f.GetCellFormula(sheet, cell)
	if err != nil {
		return false, err
	}
	if currentFormula != snapshot.Formula {
		return false, nil
	}
	if snapshot.Formula != "" {
		currentValue, err := f.GetCellValue(sheet, cell, Options{RawCellValue: true})
		if err != nil {
			return false, err
		}
		return currentValue == snapshot.Value, nil
	}
	currentValue, err := f.GetCellValue(sheet, cell, Options{RawCellValue: true})
	if err != nil {
		return false, err
	}
	currentType, err := f.GetCellType(sheet, cell)
	if err != nil {
		return false, err
	}
	return currentValue == snapshot.Value && pgMirrorValueTypeFromCell(currentValue, currentType) == snapshot.ValueType, nil
}

func (f *File) renderPostgresSnapshotValue(snapshot PGCellSnapshot) string {
	switch snapshot.ValueType {
	case pgMirrorValueTypeBool:
		if strings.EqualFold(snapshot.Value, "true") || snapshot.Value == "1" {
			return "TRUE"
		}
		return "FALSE"
	case pgMirrorValueTypeBlank:
		return ""
	default:
		return snapshot.Value
	}
}

func (f *File) applyPostgresCellSnapshot(sheet, cell string, snapshot PGCellSnapshot) error {
	if snapshot.Formula != "" {
		return f.SetCellFormula(sheet, cell, snapshot.Formula)
	}
	switch snapshot.ValueType {
	case pgMirrorValueTypeBlank:
		return f.SetCellDefault(sheet, cell, "")
	case pgMirrorValueTypeBool:
		value := strings.EqualFold(snapshot.Value, "true") || snapshot.Value == "1"
		return f.SetCellBool(sheet, cell, value)
	case pgMirrorValueTypeNumber:
		if num, err := strconv.ParseFloat(snapshot.Value, 64); err == nil {
			return f.SetCellFloat(sheet, cell, num, -1, 64)
		}
		return f.SetCellDefault(sheet, cell, snapshot.Value)
	case pgMirrorValueTypeError:
		return f.setCellErrorValue(sheet, cell, snapshot.Value)
	default:
		return f.SetCellStr(sheet, cell, snapshot.Value)
	}
}

func (f *File) setCellErrorValue(sheet, cell, value string) error {
	f.mu.Lock()
	ws, err := f.workSheetReader(sheet)
	if err != nil {
		f.mu.Unlock()
		return err
	}
	f.mu.Unlock()
	ws.mu.Lock()
	defer ws.mu.Unlock()

	c, col, row, err := ws.prepareCell(cell)
	if err != nil {
		return err
	}
	c.S = ws.prepareCellStyle(col, row, c.S)
	c.T, c.V, c.IS = "e", value, nil
	if err := f.removeFormula(c, ws, sheet); err != nil {
		return err
	}
	f.markCellValueDirty(sheet, cell)
	f.bumpSheetVersion(sheet)
	f.setWorksheetCacheCell(sheet, cell, newErrorFormulaArg(value, value))
	return nil
}

func (f *File) syncCellToPostgresAfterChange(sheet, cell string) error {
	if f.queueCellToPostgres(sheet, cell) {
		return nil
	}
	f.mu.Lock()
	enabled := f.pgMirrorEnabled
	bestEffort := f.pgMirrorBestEffort
	db := f.pgMirrorDB
	cfg := clonePGSyncOptions(f.pgMirrorOpts)
	f.mu.Unlock()

	if !enabled || db == nil || cfg == nil {
		return nil
	}

	f.pgMirrorSyncMu.Lock()
	defer f.pgMirrorSyncMu.Unlock()

	err := f.syncCellToPostgresFastPath(context.Background(), db, sheet, cell, *cfg)
	if err != nil {
		err = f.SyncCellToPostgres(context.Background(), db, sheet, cell, *cfg)
	}
	if err != nil && bestEffort {
		return nil
	}
	return err
}

func (f *File) syncCellsToPostgresAfterChange(cells []pgMirrorCellRef) error {
	f.mu.Lock()
	enabled := f.pgMirrorEnabled
	bestEffort := f.pgMirrorBestEffort
	db := f.pgMirrorDB
	cfg := clonePGSyncOptions(f.pgMirrorOpts)
	f.mu.Unlock()

	if !enabled || db == nil || cfg == nil || len(cells) == 0 {
		return nil
	}

	f.pgMirrorSyncMu.Lock()
	defer f.pgMirrorSyncMu.Unlock()

	err := f.syncCellsToPostgresFastPath(context.Background(), db, cells, *cfg)
	if err != nil {
		err = nil
		for _, ref := range cells {
			if syncErr := f.SyncCellToPostgres(context.Background(), db, ref.Sheet, ref.Cell, *cfg); syncErr != nil {
				err = syncErr
				break
			}
		}
	}
	if err != nil && bestEffort {
		return nil
	}
	return err
}

func (f *File) syncPendingCellsToPostgresAfterChange(cells []pgMirrorPendingCellRef) error {
	f.mu.Lock()
	enabled := f.pgMirrorEnabled
	bestEffort := f.pgMirrorBestEffort
	db := f.pgMirrorDB
	cfg := clonePGSyncOptions(f.pgMirrorOpts)
	f.mu.Unlock()

	if !enabled || db == nil || cfg == nil || len(cells) == 0 {
		return nil
	}

	f.pgMirrorSyncMu.Lock()
	defer f.pgMirrorSyncMu.Unlock()

	err := f.syncPendingCellsToPostgresFastPath(context.Background(), db, cells, *cfg)
	if err != nil {
		err = nil
		for _, ref := range cells {
			if syncErr := f.SyncCellToPostgres(context.Background(), db, ref.Sheet, ref.Cell, *cfg); syncErr != nil {
				err = syncErr
				break
			}
		}
	}
	if err != nil && bestEffort {
		return nil
	}
	return err
}

func (f *File) syncCellToPostgresFastPath(ctx context.Context, db *sql.DB, sheet, cell string, cfg PGSyncOptions) error {
	if !f.canUsePGMirrorFastPath() {
		return errors.New("postgres mirror fast path not initialized")
	}
	if db == nil {
		return errors.New("postgres mirror requires non-nil *sql.DB")
	}
	if err := checkSheetName(sheet); err != nil {
		return err
	}
	col, row, err := CellNameToCoordinates(cell)
	if err != nil {
		return err
	}
	if ctx == nil {
		ctx = context.Background()
	}

	value, err := f.GetCellValue(sheet, cell, Options{RawCellValue: cfg.RawCellValue})
	if err != nil {
		return err
	}
	formula, err := f.GetCellFormula(sheet, cell)
	if err != nil {
		return err
	}
	cellType, err := f.GetCellType(sheet, cell)
	if err != nil {
		return err
	}
	valueType := pgMirrorValueTypeFromCell(value, cellType)
	needsSheetMeta := f.needsPGMirrorSheetMetaUpdate(sheet, row, col)
	sqlStr, args := buildPGSingleCellFastPathSQL(cfg.Schema, getPGMirrorTables(cfg.TablePrefix), cfg.WorkbookID, sheet, pgMirrorCell{
		Cell:      cell,
		Row:       row,
		Col:       col,
		Value:     value,
		ValueType: valueType,
		Formula:   formula,
	}, needsSheetMeta)
	if _, err = db.ExecContext(ctx, sqlStr, args...); err != nil {
		return err
	}
	if needsSheetMeta {
		f.updatePGMirrorSheetMeta(sheet, row, col)
	}
	return nil
}

func (f *File) syncCellsToPostgresFastPath(ctx context.Context, db *sql.DB, refs []pgMirrorCellRef, cfg PGSyncOptions) error {
	if !f.canUsePGMirrorFastPath() {
		return errors.New("postgres mirror fast path not initialized")
	}
	if db == nil {
		return errors.New("postgres mirror requires non-nil *sql.DB")
	}
	if ctx == nil {
		ctx = context.Background()
	}

	grouped := make(map[string][]pgMirrorCell, len(refs))
	sheetMax := make(map[string]pgMirrorSheetMeta, len(refs))
	seen := make(map[string]struct{}, len(refs))
	for _, ref := range refs {
		if err := checkSheetName(ref.Sheet); err != nil {
			return err
		}
		if ref.Cell == "" {
			continue
		}
		key := ref.Sheet + "!" + ref.Cell
		if _, ok := seen[key]; ok {
			continue
		}
		seen[key] = struct{}{}

		col, row, err := CellNameToCoordinates(ref.Cell)
		if err != nil {
			return err
		}
		value, err := f.GetCellValue(ref.Sheet, ref.Cell, Options{RawCellValue: cfg.RawCellValue})
		if err != nil {
			return err
		}
		formula, err := f.GetCellFormula(ref.Sheet, ref.Cell)
		if err != nil {
			return err
		}
		cellType, err := f.GetCellType(ref.Sheet, ref.Cell)
		if err != nil {
			return err
		}
		grouped[ref.Sheet] = append(grouped[ref.Sheet], pgMirrorCell{
			Cell:      ref.Cell,
			Row:       row,
			Col:       col,
			Value:     value,
			ValueType: pgMirrorValueTypeFromCell(value, cellType),
			Formula:   formula,
		})
		meta := sheetMax[ref.Sheet]
		if row > meta.RowCount {
			meta.RowCount = row
		}
		if col > meta.ColCount {
			meta.ColCount = col
		}
		sheetMax[ref.Sheet] = meta
	}
	if len(grouped) == 0 {
		return nil
	}

	tx, err := db.BeginTx(ctx, nil)
	if err != nil {
		return err
	}
	rollback := true
	defer func() {
		if rollback {
			_ = tx.Rollback()
		}
	}()

	tables := getPGMirrorTables(cfg.TablePrefix)
	for sheet, cells := range grouped {
		meta := sheetMax[sheet]
		if f.needsPGMirrorSheetMetaUpdate(sheet, meta.RowCount, meta.ColCount) {
			if err := upsertPGSheetMirror(ctx, tx, cfg.Schema, tables, cfg.WorkbookID, sheet, meta.RowCount, meta.ColCount, -1); err != nil {
				return err
			}
		}

		upserts := make([]pgMirrorCell, 0, len(cells))
		deletes := make([]string, 0, len(cells))
		for _, cell := range cells {
			if cell.Value == "" && cell.Formula == "" {
				deletes = append(deletes, cell.Cell)
				continue
			}
			upserts = append(upserts, cell)
		}
		if len(deletes) > 0 {
			if err := deletePGMirroredCells(ctx, tx, cfg.Schema, tables, cfg.WorkbookID, sheet, deletes); err != nil {
				return err
			}
			if err := deletePGLookupRows(ctx, tx, cfg.Schema, tables, cfg.WorkbookID, sheet, deletes); err != nil {
				return err
			}
		}
		if len(upserts) > 0 {
			if err := upsertPGMirroredCells(ctx, tx, cfg.Schema, tables, cfg.WorkbookID, sheet, upserts); err != nil {
				return err
			}
			if err := upsertPGLookupRows(ctx, tx, cfg.Schema, tables, cfg.WorkbookID, sheet, upserts); err != nil {
				return err
			}
		}
	}

	if err := tx.Commit(); err != nil {
		return err
	}
	rollback = false

	for sheet, meta := range sheetMax {
		if meta.RowCount > 0 || meta.ColCount > 0 {
			f.updatePGMirrorSheetMeta(sheet, meta.RowCount, meta.ColCount)
		}
	}
	return nil
}

func (f *File) syncPendingCellsToPostgresFastPath(ctx context.Context, db *sql.DB, refs []pgMirrorPendingCellRef, cfg PGSyncOptions) error {
	if !f.canUsePGMirrorFastPath() {
		return errors.New("postgres mirror fast path not initialized")
	}
	if db == nil {
		return errors.New("postgres mirror requires non-nil *sql.DB")
	}
	if ctx == nil {
		ctx = context.Background()
	}

	grouped := make(map[string][]pgMirrorCell, len(refs))
	sheetMax := make(map[string]pgMirrorSheetMeta, len(refs))
	seen := make(map[string]struct{}, len(refs))
	for _, ref := range refs {
		if err := checkSheetName(ref.Sheet); err != nil {
			return err
		}
		if ref.Cell == "" {
			continue
		}
		key := ref.Sheet + "!" + ref.Cell
		if _, ok := seen[key]; ok {
			continue
		}
		seen[key] = struct{}{}

		col, row, err := CellNameToCoordinates(ref.Cell)
		if err != nil {
			return err
		}
		formula, err := f.GetCellFormula(ref.Sheet, ref.Cell)
		if err != nil {
			return err
		}

		value := ref.Value
		valueType := ref.ValueType
		if !ref.HasValue {
			value, err = f.GetCellValue(ref.Sheet, ref.Cell, Options{RawCellValue: cfg.RawCellValue})
			if err != nil {
				return err
			}
			cellType, err := f.GetCellType(ref.Sheet, ref.Cell)
			if err != nil {
				return err
			}
			valueType = pgMirrorValueTypeFromCell(value, cellType)
		}

		grouped[ref.Sheet] = append(grouped[ref.Sheet], pgMirrorCell{
			Cell:      ref.Cell,
			Row:       row,
			Col:       col,
			Value:     value,
			ValueType: valueType,
			Formula:   formula,
		})
		meta := sheetMax[ref.Sheet]
		if row > meta.RowCount {
			meta.RowCount = row
		}
		if col > meta.ColCount {
			meta.ColCount = col
		}
		sheetMax[ref.Sheet] = meta
	}
	if len(grouped) == 0 {
		return nil
	}

	tx, err := db.BeginTx(ctx, nil)
	if err != nil {
		return err
	}
	rollback := true
	defer func() {
		if rollback {
			_ = tx.Rollback()
		}
	}()

	tables := getPGMirrorTables(cfg.TablePrefix)
	for sheet, cells := range grouped {
		meta := sheetMax[sheet]
		if f.needsPGMirrorSheetMetaUpdate(sheet, meta.RowCount, meta.ColCount) {
			if err := upsertPGSheetMirror(ctx, tx, cfg.Schema, tables, cfg.WorkbookID, sheet, meta.RowCount, meta.ColCount, -1); err != nil {
				return err
			}
		}

		upserts := make([]pgMirrorCell, 0, len(cells))
		deletes := make([]string, 0, len(cells))
		for _, cell := range cells {
			if cell.Value == "" && cell.Formula == "" {
				deletes = append(deletes, cell.Cell)
				continue
			}
			upserts = append(upserts, cell)
		}
		if len(deletes) > 0 {
			if err := deletePGMirroredCells(ctx, tx, cfg.Schema, tables, cfg.WorkbookID, sheet, deletes); err != nil {
				return err
			}
			if err := deletePGLookupRows(ctx, tx, cfg.Schema, tables, cfg.WorkbookID, sheet, deletes); err != nil {
				return err
			}
		}
		if len(upserts) > 0 {
			if err := upsertPGMirroredCells(ctx, tx, cfg.Schema, tables, cfg.WorkbookID, sheet, upserts); err != nil {
				return err
			}
			if err := upsertPGLookupRows(ctx, tx, cfg.Schema, tables, cfg.WorkbookID, sheet, upserts); err != nil {
				return err
			}
		}
	}

	if err := tx.Commit(); err != nil {
		return err
	}
	rollback = false

	for sheet, meta := range sheetMax {
		if meta.RowCount > 0 || meta.ColCount > 0 {
			f.updatePGMirrorSheetMeta(sheet, meta.RowCount, meta.ColCount)
		}
	}
	return nil
}

func (f *File) syncCalculatedCellToPostgres(sheet, cell string) {
	if err := f.syncCellToPostgresAfterChange(sheet, cell); err != nil {
		log.Printf("excelize: postgres mirror sync failed for %s!%s: %v", sheet, cell, err)
	}
}

func (f *File) queueCellToPostgres(sheet, cell string) bool {
	if sheet == "" || cell == "" {
		return false
	}
	f.mu.Lock()
	defer f.mu.Unlock()
	if !f.pgMirrorEnabled || f.pgMirrorDB == nil || f.pgMirrorOpts == nil || f.pgMirrorBatchDepth == 0 {
		return false
	}
	if f.pgMirrorPending == nil {
		f.pgMirrorPending = make(map[string]map[string]pgMirrorPendingCell)
	}
	if f.pgMirrorPending[sheet] == nil {
		f.pgMirrorPending[sheet] = make(map[string]pgMirrorPendingCell)
	}
	if _, ok := f.pgMirrorPending[sheet][cell]; !ok {
		f.pgMirrorPending[sheet][cell] = pgMirrorPendingCell{}
	}
	return true
}

func (f *File) queueCalculatedCellToPostgres(sheet, cell, value string) bool {
	if sheet == "" || cell == "" {
		return false
	}
	f.mu.Lock()
	defer f.mu.Unlock()
	if !f.pgMirrorEnabled || f.pgMirrorDB == nil || f.pgMirrorOpts == nil || f.pgMirrorBatchDepth == 0 {
		return false
	}
	if f.pgMirrorPending == nil {
		f.pgMirrorPending = make(map[string]map[string]pgMirrorPendingCell)
	}
	if f.pgMirrorPending[sheet] == nil {
		f.pgMirrorPending[sheet] = make(map[string]pgMirrorPendingCell)
	}
	f.pgMirrorPending[sheet][cell] = pgMirrorPendingCell{
		HasValue:  true,
		Value:     value,
		ValueType: pgMirrorValueTypeFromCell(value, xmlCellTypeToCellType(inferXMLCellType(value))),
	}
	return true
}

func (f *File) beginPGMirrorCalculationBatch() {
	f.mu.Lock()
	defer f.mu.Unlock()
	if !f.pgMirrorEnabled || f.pgMirrorDB == nil || f.pgMirrorOpts == nil {
		return
	}
	f.pgMirrorBatchDepth++
	if f.pgMirrorPending == nil {
		f.pgMirrorPending = make(map[string]map[string]pgMirrorPendingCell)
		return
	}
	if f.pgMirrorBatchDepth == 1 {
		for sheet := range f.pgMirrorPending {
			delete(f.pgMirrorPending, sheet)
		}
	}
}

func (f *File) flushPGMirrorCalculationBatch() error {
	f.mu.Lock()
	if f.pgMirrorBatchDepth == 0 {
		f.mu.Unlock()
		return nil
	}
	f.pgMirrorBatchDepth--
	if f.pgMirrorBatchDepth > 0 {
		f.mu.Unlock()
		return nil
	}
	pending := make([]pgMirrorPendingCellRef, 0)
	for sheet, cells := range f.pgMirrorPending {
		for cell, pendingCell := range cells {
			pending = append(pending, pgMirrorPendingCellRef{
				Sheet:     sheet,
				Cell:      cell,
				HasValue:  pendingCell.HasValue,
				Value:     pendingCell.Value,
				ValueType: pendingCell.ValueType,
			})
		}
	}
	if f.pgMirrorPending == nil {
		f.pgMirrorPending = make(map[string]map[string]pgMirrorPendingCell)
	} else {
		for sheet := range f.pgMirrorPending {
			delete(f.pgMirrorPending, sheet)
		}
	}
	f.mu.Unlock()

	if len(pending) == 0 {
		return nil
	}
	return f.syncPendingCellsToPostgresAfterChange(pending)
}

func (f *File) notifyCellCalculated(sheet, cell, oldValue, newValue string) {
	if cell == "" || oldValue == newValue {
		return
	}
	if f.OnCellCalculated != nil {
		f.OnCellCalculated(sheet, cell, oldValue, newValue)
	}
	if f.queueCalculatedCellToPostgres(sheet, cell, newValue) {
		return
	}
	f.syncCalculatedCellToPostgres(sheet, cell)
}

func xmlCellTypeToCellType(xmlType string) CellType {
	switch xmlType {
	case "b":
		return CellTypeBool
	case "e":
		return CellTypeError
	case "str":
		return CellTypeFormula
	default:
		return CellTypeUnset
	}
}

func (f *File) syncSheetToPostgresTx(ctx context.Context, tx *sql.Tx, cfg *PGSyncOptions, tables pgMirrorTables, sheet string) (pgMirrorSheetMeta, error) {
	cells, rowCount, colCount, err := f.collectSheetMirrorCells(sheet, cfg.RawCellValue)
	if err != nil {
		return pgMirrorSheetMeta{}, err
	}
	if err = upsertPGSheetMirror(ctx, tx, cfg.Schema, tables, cfg.WorkbookID, sheet, rowCount, colCount, len(cells)); err != nil {
		return pgMirrorSheetMeta{}, err
	}
	if err = replacePGMirroredSheetCells(ctx, tx, cfg.Schema, tables, cfg.WorkbookID, sheet, cells, cfg.BatchSize); err != nil {
		return pgMirrorSheetMeta{}, err
	}
	if err = replacePGMirroredSheetLookups(ctx, tx, cfg.Schema, tables, cfg.WorkbookID, sheet, cells, cfg.BatchSize); err != nil {
		return pgMirrorSheetMeta{}, err
	}
	return pgMirrorSheetMeta{RowCount: rowCount, ColCount: colCount}, nil
}

func (f *File) trySyncWorkbookToPostgresViaPGXCopy(ctx context.Context, db *sql.DB, cfg *PGSyncOptions) (map[string]pgMirrorSheetMeta, bool, error) {
	conn, isPGX, err := openPGXSQLConn(ctx, db)
	if err != nil {
		return nil, true, err
	}
	if !isPGX {
		return nil, false, nil
	}
	defer conn.Close()

	var sheetMeta map[string]pgMirrorSheetMeta
	if err := execPGMirrorTx(ctx, conn, func(exec pgMirrorExecutor) error {
		tables := getPGMirrorTables(cfg.TablePrefix)
		sheets := f.GetSheetList()
		sheetMeta = make(map[string]pgMirrorSheetMeta, len(sheets))
		if err := ensurePGMirrorTables(ctx, exec, cfg.Schema, tables); err != nil {
			return err
		}
		if err := upsertPGWorkbookMirror(ctx, exec, cfg.Schema, tables, cfg.WorkbookID, f.Path); err != nil {
			return err
		}
		stageTable := buildPGMirrorStageTableName(tables)
		if err := ensurePGMirrorStageTable(ctx, exec, stageTable); err != nil {
			return err
		}
		for _, sheet := range sheets {
			meta, err := f.syncSheetToPostgresViaPGXCopy(ctx, conn, exec, cfg, tables, stageTable, sheet)
			if err != nil {
				return err
			}
			sheetMeta[sheet] = meta
		}
		if err := deletePGRemovedSheets(ctx, exec, cfg.Schema, tables, cfg.WorkbookID, sheets); err != nil {
			return err
		}
		return nil
	}); err != nil {
		return nil, true, err
	}

	return sheetMeta, true, nil
}

func (f *File) trySyncSheetToPostgresViaPGXCopy(ctx context.Context, db *sql.DB, sheet string, cfg *PGSyncOptions) (pgMirrorSheetMeta, bool, error) {
	conn, isPGX, err := openPGXSQLConn(ctx, db)
	if err != nil {
		return pgMirrorSheetMeta{}, true, err
	}
	if !isPGX {
		return pgMirrorSheetMeta{}, false, nil
	}
	defer conn.Close()

	var meta pgMirrorSheetMeta
	if err := execPGMirrorTx(ctx, conn, func(exec pgMirrorExecutor) error {
		tables := getPGMirrorTables(cfg.TablePrefix)
		if err := ensurePGMirrorTables(ctx, exec, cfg.Schema, tables); err != nil {
			return err
		}
		if err := upsertPGWorkbookMirror(ctx, exec, cfg.Schema, tables, cfg.WorkbookID, f.Path); err != nil {
			return err
		}
		stageTable := buildPGMirrorStageTableName(tables)
		if err := ensurePGMirrorStageTable(ctx, exec, stageTable); err != nil {
			return err
		}
		var err error
		meta, err = f.syncSheetToPostgresViaPGXCopy(ctx, conn, exec, cfg, tables, stageTable, sheet)
		return err
	}); err != nil {
		return pgMirrorSheetMeta{}, true, err
	}
	return meta, true, nil
}

func (f *File) syncSheetToPostgresViaPGXCopy(ctx context.Context, conn *sql.Conn, exec pgMirrorExecutor, cfg *PGSyncOptions, tables pgMirrorTables, stageTable, sheet string) (pgMirrorSheetMeta, error) {
	cells, rowCount, colCount, err := f.collectSheetMirrorCells(sheet, cfg.RawCellValue)
	if err != nil {
		return pgMirrorSheetMeta{}, err
	}
	if err = upsertPGSheetMirror(ctx, exec, cfg.Schema, tables, cfg.WorkbookID, sheet, rowCount, colCount, len(cells)); err != nil {
		return pgMirrorSheetMeta{}, err
	}
	if err = replacePGMirroredSheetViaPGXCopy(ctx, conn, exec, cfg.Schema, tables, cfg.WorkbookID, sheet, stageTable, cells); err != nil {
		return pgMirrorSheetMeta{}, err
	}
	return pgMirrorSheetMeta{RowCount: rowCount, ColCount: colCount}, nil
}

func openPGXSQLConn(ctx context.Context, db *sql.DB) (*sql.Conn, bool, error) {
	conn, err := db.Conn(ctx)
	if err != nil {
		return nil, false, err
	}
	isPGX := false
	err = conn.Raw(func(driverConn any) error {
		_, isPGX = driverConn.(*stdlib.Conn)
		return nil
	})
	if err != nil {
		_ = conn.Close()
		return nil, false, err
	}
	if !isPGX {
		_ = conn.Close()
		return nil, false, nil
	}
	return conn, true, nil
}

func execPGMirrorTx(ctx context.Context, conn *sql.Conn, fn func(pgMirrorExecutor) error) (err error) {
	if _, err = conn.ExecContext(ctx, "BEGIN"); err != nil {
		return err
	}
	committed := false
	defer func() {
		if committed {
			return
		}
		_, _ = conn.ExecContext(ctx, "ROLLBACK")
	}()
	if err = fn(conn); err != nil {
		return err
	}
	if _, err = conn.ExecContext(ctx, "COMMIT"); err != nil {
		return err
	}
	committed = true
	return nil
}

func ensurePGMirrorStageTable(ctx context.Context, exec pgMirrorExecutor, stageTable string) error {
	sqlStr := fmt.Sprintf(`
CREATE TEMP TABLE IF NOT EXISTS %s (
	cell_ref TEXT NOT NULL,
	row_num INTEGER NOT NULL,
	col_num INTEGER NOT NULL,
	value TEXT NOT NULL DEFAULT '',
	value_type TEXT NOT NULL DEFAULT 'blank',
	formula TEXT NOT NULL DEFAULT ''
) ON COMMIT DROP`, quotePGIdent(stageTable))
	_, err := exec.ExecContext(ctx, sqlStr)
	return err
}

func buildPGMirrorStageTableName(tables pgMirrorTables) string {
	return tables.cell + "_stage"
}

func replacePGMirroredSheetViaPGXCopy(ctx context.Context, conn *sql.Conn, exec pgMirrorExecutor, schema string, tables pgMirrorTables, workbookID, sheet, stageTable string, cells []pgMirrorCell) error {
	if _, err := exec.ExecContext(ctx,
		fmt.Sprintf(`DELETE FROM %s WHERE workbook_id = $1 AND sheet_name = $2`, qualifyPGTable(schema, tables.cell)),
		workbookID, sheet,
	); err != nil {
		return err
	}
	if _, err := exec.ExecContext(ctx,
		fmt.Sprintf(`DELETE FROM %s WHERE workbook_id = $1 AND sheet_name = $2`, qualifyPGTable(schema, tables.lookup)),
		workbookID, sheet,
	); err != nil {
		return err
	}
	if len(cells) == 0 {
		return nil
	}
	if _, err := exec.ExecContext(ctx, `TRUNCATE TABLE `+quotePGIdent(stageTable)); err != nil {
		return err
	}
	if err := copyPGMirrorCellsToStage(ctx, conn, stageTable, cells); err != nil {
		return err
	}
	if _, err := exec.ExecContext(ctx, buildPGInsertMirroredSheetFromStageSQL(schema, tables, stageTable), workbookID, sheet); err != nil {
		return err
	}
	if pgMirrorCopyFromTestHook != nil {
		pgMirrorCopyFromTestHook(sheet, len(cells))
	}
	return nil
}

func copyPGMirrorCellsToStage(ctx context.Context, conn *sql.Conn, stageTable string, cells []pgMirrorCell) error {
	return conn.Raw(func(driverConn any) error {
		stdConn, ok := driverConn.(*stdlib.Conn)
		if !ok {
			return errors.New("postgres mirror copy path requires pgx stdlib connection")
		}
		_, err := stdConn.Conn().CopyFrom(
			ctx,
			pgx.Identifier{"pg_temp", stageTable},
			[]string{"cell_ref", "row_num", "col_num", "value", "value_type", "formula"},
			pgx.CopyFromSlice(len(cells), func(i int) ([]any, error) {
				record := cells[i]
				return []any{
					record.Cell,
					int32(record.Row),
					int32(record.Col),
					record.Value,
					record.ValueType,
					record.Formula,
				}, nil
			}),
		)
		return err
	})
}

func buildPGInsertMirroredSheetFromStageSQL(schema string, tables pgMirrorTables, stageTable string) string {
	return fmt.Sprintf(`
WITH inserted AS (
	INSERT INTO %s(workbook_id, sheet_name, cell_ref, row_num, col_num, value, value_type, formula, updated_at)
	SELECT $1, $2, cell_ref, row_num, col_num, value, value_type, formula, NOW()
	FROM %s
	RETURNING cell_ref, row_num, col_num, value, value_type, updated_at
)
INSERT INTO %s(workbook_id, sheet_name, cell_ref, row_num, col_num, cell_value, cell_value_type, updated_at)
SELECT $1, $2, cell_ref, row_num, col_num, value, value_type, updated_at
FROM inserted`,
		qualifyPGTable(schema, tables.cell),
		quotePGIdent(stageTable),
		qualifyPGTable(schema, tables.lookup),
	)
}

func (f *File) collectSheetMirrorCells(sheet string, rawCellValue bool) ([]pgMirrorCell, int, int, error) {
	ws, err := f.workSheetReader(sheet)
	if err != nil {
		return nil, 0, 0, err
	}
	if ws == nil || ws.SheetData.Row == nil {
		return nil, 0, 0, nil
	}

	sst, err := f.sharedStringsReader()
	if err != nil {
		return nil, 0, 0, err
	}

	maxRow, maxCol := 0, 0
	cells := make([]pgMirrorCell, 0, len(ws.SheetData.Row)*2)
	for rowIdx, row := range ws.SheetData.Row {
		rowNum := row.R
		if rowNum == 0 {
			rowNum = rowIdx + 1
		}
		if rowNum > maxRow {
			maxRow = rowNum
		}

		for colIdx := range row.C {
			cell := &row.C[colIdx]
			cellRef := cell.R
			colNum, parsedRowNum, refErr := CellNameToCoordinates(cellRef)
			if refErr != nil {
				colNum = colIdx + 1
				parsedRowNum = rowNum
				cellRef, refErr = CoordinatesToCellName(colNum, parsedRowNum)
				if refErr != nil {
					return nil, 0, 0, refErr
				}
			}
			if parsedRowNum > maxRow {
				maxRow = parsedRowNum
			}
			if colNum > maxCol {
				maxCol = colNum
			}

			value, valueErr := cell.getValueFrom(f, sst, rawCellValue)
			if valueErr != nil {
				return nil, 0, 0, valueErr
			}
			formula := ""
			if cell.F != nil {
				formula = cell.F.Content
				if formula == "" && cell.F.T == STCellFormulaTypeShared && cell.F.Si != nil {
					formula, err = getSharedFormula(ws, *cell.F.Si, cellRef)
					if err != nil {
						return nil, 0, 0, err
					}
				}
			}
			if value == "" && formula == "" {
				continue
			}

			cells = append(cells, pgMirrorCell{
				Cell:      cellRef,
				Row:       parsedRowNum,
				Col:       colNum,
				Value:     value,
				ValueType: pgMirrorValueTypeFromCell(value, cellTypes[cell.T]),
				Formula:   formula,
			})
		}
	}
	return cells, maxRow, maxCol, nil
}

func (f *File) canUsePGMirrorFastPath() bool {
	f.mu.Lock()
	defer f.mu.Unlock()
	return f.pgMirrorReady
}

func (f *File) markPGMirrorReady(ready bool) {
	f.mu.Lock()
	f.pgMirrorReady = ready
	if !ready && f.pgMirrorSheetMeta == nil {
		f.pgMirrorSheetMeta = make(map[string]pgMirrorSheetMeta)
	}
	f.mu.Unlock()
}

func (f *File) needsPGMirrorSheetMetaUpdate(sheet string, row, col int) bool {
	f.mu.Lock()
	defer f.mu.Unlock()
	meta, ok := f.pgMirrorSheetMeta[sheet]
	if !ok {
		return true
	}
	return row > meta.RowCount || col > meta.ColCount
}

func (f *File) updatePGMirrorSheetMeta(sheet string, rowCount, colCount int) {
	f.mu.Lock()
	if f.pgMirrorSheetMeta == nil {
		f.pgMirrorSheetMeta = make(map[string]pgMirrorSheetMeta)
	}
	meta := f.pgMirrorSheetMeta[sheet]
	if rowCount > meta.RowCount {
		meta.RowCount = rowCount
	}
	if colCount > meta.ColCount {
		meta.ColCount = colCount
	}
	f.pgMirrorSheetMeta[sheet] = meta
	f.pgMirrorReady = true
	f.mu.Unlock()
}

func (f *File) refreshPGMirrorSheetMeta(rawCellValue bool) {
	sheets := f.GetSheetList()
	meta := make(map[string]pgMirrorSheetMeta, len(sheets))
	for _, sheet := range sheets {
		_, rowCount, colCount, err := f.collectSheetMirrorCells(sheet, rawCellValue)
		if err != nil {
			continue
		}
		meta[sheet] = pgMirrorSheetMeta{RowCount: rowCount, ColCount: colCount}
	}
	f.mu.Lock()
	f.pgMirrorSheetMeta = meta
	f.pgMirrorReady = true
	f.mu.Unlock()
}

func (f *File) refreshPGMirrorSheetMetaForSheet(sheet string, rawCellValue bool) {
	_, rowCount, colCount, err := f.collectSheetMirrorCells(sheet, rawCellValue)
	if err != nil {
		return
	}
	f.updatePGMirrorSheetMeta(sheet, rowCount, colCount)
}

func (f *File) replacePGMirrorSheetMeta(meta map[string]pgMirrorSheetMeta) {
	f.mu.Lock()
	if f.pgMirrorSheetMeta == nil {
		f.pgMirrorSheetMeta = make(map[string]pgMirrorSheetMeta, len(meta))
	} else {
		for sheet := range f.pgMirrorSheetMeta {
			delete(f.pgMirrorSheetMeta, sheet)
		}
	}
	for sheet, sheetMeta := range meta {
		f.pgMirrorSheetMeta[sheet] = sheetMeta
	}
	f.pgMirrorReady = true
	f.mu.Unlock()
}

func (f *File) mergePGMirrorSheetMeta(meta map[string]pgMirrorSheetMeta) {
	f.mu.Lock()
	if f.pgMirrorSheetMeta == nil {
		f.pgMirrorSheetMeta = make(map[string]pgMirrorSheetMeta, len(meta))
	}
	for sheet, sheetMeta := range meta {
		f.pgMirrorSheetMeta[sheet] = sheetMeta
	}
	f.pgMirrorReady = true
	f.mu.Unlock()
}

func (f *File) getPGSyncOptions(opts ...PGSyncOptions) (*PGSyncOptions, error) {
	cfg := &PGSyncOptions{
		Schema:      "public",
		TablePrefix: "excelize",
		BatchSize:   1000,
	}
	for _, opt := range opts {
		cfg = &opt
	}
	if cfg.Schema == "" {
		cfg.Schema = "public"
	}
	if cfg.TablePrefix == "" {
		cfg.TablePrefix = "excelize"
	}
	if cfg.BatchSize <= 0 {
		cfg.BatchSize = 1000
	}
	if !isValidPGIdent(cfg.Schema) {
		return nil, fmt.Errorf("invalid postgres schema %q", cfg.Schema)
	}
	if !isValidPGIdent(cfg.TablePrefix) {
		return nil, fmt.Errorf("invalid postgres table prefix %q", cfg.TablePrefix)
	}
	if cfg.WorkbookID == "" {
		cfg.WorkbookID = deriveWorkbookID(f.Path)
	}
	return clonePGSyncOptions(cfg), nil
}

func clonePGSyncOptions(cfg *PGSyncOptions) *PGSyncOptions {
	if cfg == nil {
		return nil
	}
	cloned := *cfg
	if cfg.WarmupSheets != nil {
		cloned.WarmupSheets = append([]string(nil), cfg.WarmupSheets...)
	}
	return &cloned
}

func clonePGWarmupStatus(status PGWarmupStatus) PGWarmupStatus {
	if status.Sheets != nil {
		status.Sheets = append([]string(nil), status.Sheets...)
	}
	return status
}

func (f *File) resetPGWarmupLocked(enabled bool) {
	if f.pgWarmupCancel != nil {
		f.pgWarmupCancel()
		f.pgWarmupCancel = nil
	}
	if f.pgWarmupDone != nil {
		close(f.pgWarmupDone)
		f.pgWarmupDone = nil
	}
	f.pgWarmupSeq++
	f.pgWarmupStatus = PGWarmupStatus{Enabled: enabled}
}

func (f *File) startPGWarmupLocked(cfg *PGSyncOptions) (uint64, context.Context) {
	f.resetPGWarmupLocked(true)

	warmupCtx := context.Background()
	var cancel context.CancelFunc
	if cfg != nil && cfg.WarmupTimeout > 0 {
		warmupCtx, cancel = context.WithTimeout(warmupCtx, cfg.WarmupTimeout)
	} else {
		warmupCtx, cancel = context.WithCancel(warmupCtx)
	}

	id := f.pgWarmupSeq + 1
	f.pgWarmupSeq = id
	f.pgWarmupCancel = cancel
	f.pgWarmupDone = make(chan struct{})
	f.pgWarmupStatus = PGWarmupStatus{
		Enabled:    true,
		InProgress: true,
		Async:      cfg != nil && cfg.AsyncWarmup,
		StartedAt:  time.Now(),
		Err:        nil,
	}
	if cfg != nil && cfg.WarmupSheets != nil {
		f.pgWarmupStatus.Sheets = append([]string(nil), cfg.WarmupSheets...)
	}
	return id, warmupCtx
}

func (f *File) preloadPostgresWarmup(ctx context.Context, cfg *PGSyncOptions) error {
	if pgMirrorWarmupTestHook != nil {
		pgMirrorWarmupTestHook()
	}
	if cfg == nil {
		return nil
	}
	return f.PreloadPostgresLookupCache(ctx, cfg.WarmupSheets...)
}

func (f *File) runPostgresWarmup(id uint64, ctx context.Context, cfg *PGSyncOptions) {
	err := f.preloadPostgresWarmup(ctx, cfg)
	f.finishPGWarmup(id, err)
	if err != nil && !errors.Is(err, context.Canceled) {
		log.Printf("excelize: postgres mirror auto warmup failed: %v", err)
	}
}

func (f *File) finishPGWarmup(id uint64, err error) {
	f.mu.Lock()
	defer f.mu.Unlock()

	if id != f.pgWarmupSeq {
		return
	}
	status := f.pgWarmupStatus
	status.Enabled = f.pgMirrorEnabled && f.pgMirrorDB != nil && f.pgMirrorOpts != nil
	status.InProgress = false
	status.CompletedAt = time.Now()
	status.Err = err
	f.pgWarmupStatus = clonePGWarmupStatus(status)
	if f.pgWarmupCancel != nil {
		f.pgWarmupCancel()
		f.pgWarmupCancel = nil
	}
	if f.pgWarmupDone != nil {
		close(f.pgWarmupDone)
		f.pgWarmupDone = nil
	}
}

func deriveWorkbookID(path string) string {
	if path == "" {
		return "workbook"
	}
	base := filepath.Base(path)
	ext := filepath.Ext(base)
	if ext != "" {
		base = strings.TrimSuffix(base, ext)
	}
	base = strings.TrimSpace(base)
	if base == "" {
		return "workbook"
	}
	return base
}

func getPGMirrorTables(prefix string) pgMirrorTables {
	return pgMirrorTables{
		workbook: prefix + "_workbook_mirror",
		sheet:    prefix + "_sheet_mirror",
		cell:     prefix + "_cell_mirror",
		lookup:   prefix + "_lookup_mirror",
	}
}

func isValidPGIdent(name string) bool {
	return pgIdentRegexp.MatchString(name)
}

func quotePGIdent(name string) string {
	return `"` + strings.ReplaceAll(name, `"`, `""`) + `"`
}

func qualifyPGTable(schema, table string) string {
	return quotePGIdent(schema) + "." + quotePGIdent(table)
}

func ensurePGMirrorTables(ctx context.Context, exec pgMirrorExecutor, schema string, tables pgMirrorTables) error {
	sqls := []string{
		fmt.Sprintf(`CREATE SCHEMA IF NOT EXISTS %s`, quotePGIdent(schema)),
		fmt.Sprintf(`
CREATE TABLE IF NOT EXISTS %s (
	workbook_id TEXT PRIMARY KEY,
	file_path TEXT NOT NULL DEFAULT '',
	synced_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
)`, qualifyPGTable(schema, tables.workbook)),
		fmt.Sprintf(`
CREATE TABLE IF NOT EXISTS %s (
	workbook_id TEXT NOT NULL,
	sheet_name TEXT NOT NULL,
	row_count INTEGER NOT NULL DEFAULT 0,
	col_count INTEGER NOT NULL DEFAULT 0,
	cell_count INTEGER NOT NULL DEFAULT 0,
	last_synced_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
	PRIMARY KEY (workbook_id, sheet_name),
	FOREIGN KEY (workbook_id) REFERENCES %s(workbook_id) ON DELETE CASCADE
)`, qualifyPGTable(schema, tables.sheet), qualifyPGTable(schema, tables.workbook)),
		fmt.Sprintf(`
CREATE TABLE IF NOT EXISTS %s (
	workbook_id TEXT NOT NULL,
	sheet_name TEXT NOT NULL,
	cell_ref TEXT NOT NULL,
	row_num INTEGER NOT NULL,
	col_num INTEGER NOT NULL,
	value TEXT NOT NULL DEFAULT '',
	value_type TEXT NOT NULL DEFAULT 'blank',
	formula TEXT NOT NULL DEFAULT '',
	updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
	PRIMARY KEY (workbook_id, sheet_name, cell_ref),
	FOREIGN KEY (workbook_id, sheet_name) REFERENCES %s(workbook_id, sheet_name) ON DELETE CASCADE
)`, qualifyPGTable(schema, tables.cell), qualifyPGTable(schema, tables.sheet)),
		fmt.Sprintf(`
CREATE TABLE IF NOT EXISTS %s (
	workbook_id TEXT NOT NULL,
	sheet_name TEXT NOT NULL,
	cell_ref TEXT NOT NULL,
	row_num INTEGER NOT NULL,
	col_num INTEGER NOT NULL,
	cell_value TEXT NOT NULL DEFAULT '',
	cell_value_type TEXT NOT NULL DEFAULT 'blank',
	updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
	PRIMARY KEY (workbook_id, sheet_name, cell_ref),
	FOREIGN KEY (workbook_id, sheet_name) REFERENCES %s(workbook_id, sheet_name) ON DELETE CASCADE
)`, qualifyPGTable(schema, tables.lookup), qualifyPGTable(schema, tables.sheet)),
		fmt.Sprintf(`ALTER TABLE %s ADD COLUMN IF NOT EXISTS value_type TEXT NOT NULL DEFAULT 'blank'`,
			qualifyPGTable(schema, tables.cell)),
		fmt.Sprintf(`ALTER TABLE %s ADD COLUMN IF NOT EXISTS cell_value_type TEXT NOT NULL DEFAULT 'blank'`,
			qualifyPGTable(schema, tables.lookup)),
		fmt.Sprintf(`CREATE INDEX IF NOT EXISTS %s ON %s(workbook_id, sheet_name, row_num, col_num)`,
			quotePGIdent(tables.cell+"_row_col_idx"), qualifyPGTable(schema, tables.cell)),
		fmt.Sprintf(`CREATE INDEX IF NOT EXISTS %s ON %s(workbook_id, sheet_name, col_num, cell_value_type, cell_value, row_num)`,
			quotePGIdent(tables.lookup+"_col_type_value_row_idx"), qualifyPGTable(schema, tables.lookup)),
		fmt.Sprintf(`CREATE INDEX IF NOT EXISTS %s ON %s(workbook_id, sheet_name, row_num, cell_value_type, cell_value, col_num)`,
			quotePGIdent(tables.lookup+"_row_type_value_col_idx"), qualifyPGTable(schema, tables.lookup)),
	}
	for _, stmt := range sqls {
		if _, err := exec.ExecContext(ctx, stmt); err != nil {
			return err
		}
	}
	return nil
}

func upsertPGWorkbookMirror(ctx context.Context, exec pgMirrorExecutor, schema string, tables pgMirrorTables, workbookID, filePath string) error {
	sqlStr := fmt.Sprintf(`
INSERT INTO %s(workbook_id, file_path, synced_at)
VALUES ($1, $2, NOW())
ON CONFLICT (workbook_id)
DO UPDATE SET file_path = EXCLUDED.file_path, synced_at = NOW()`, qualifyPGTable(schema, tables.workbook))
	_, err := exec.ExecContext(ctx, sqlStr, workbookID, filePath)
	return err
}

func upsertPGSheetMirror(ctx context.Context, exec pgMirrorExecutor, schema string, tables pgMirrorTables, workbookID, sheet string, rowCount, colCount, cellCount int) error {
	sqlStr := fmt.Sprintf(`
INSERT INTO %s AS tgt(workbook_id, sheet_name, row_count, col_count, cell_count, last_synced_at)
VALUES ($1, $2, $3, $4, $5, NOW())
ON CONFLICT (workbook_id, sheet_name)
DO UPDATE SET
	row_count = GREATEST(tgt.row_count, EXCLUDED.row_count),
	col_count = GREATEST(tgt.col_count, EXCLUDED.col_count),
	cell_count = CASE WHEN EXCLUDED.cell_count < 0 THEN tgt.cell_count ELSE EXCLUDED.cell_count END,
	last_synced_at = NOW()`,
		qualifyPGTable(schema, tables.sheet),
	)
	_, err := exec.ExecContext(ctx, sqlStr, workbookID, sheet, rowCount, colCount, cellCount)
	return err
}

func replacePGMirroredSheetCells(ctx context.Context, exec pgMirrorExecutor, schema string, tables pgMirrorTables, workbookID, sheet string, cells []pgMirrorCell, batchSize int) error {
	deleteSQL := fmt.Sprintf(`DELETE FROM %s WHERE workbook_id = $1 AND sheet_name = $2`, qualifyPGTable(schema, tables.cell))
	if _, err := exec.ExecContext(ctx, deleteSQL, workbookID, sheet); err != nil {
		return err
	}
	if len(cells) == 0 {
		return nil
	}
	for start := 0; start < len(cells); start += batchSize {
		end := start + batchSize
		if end > len(cells) {
			end = len(cells)
		}
		if err := upsertPGMirroredCells(ctx, exec, schema, tables, workbookID, sheet, cells[start:end]); err != nil {
			return err
		}
	}
	return nil
}

func replacePGMirroredSheetLookups(ctx context.Context, exec pgMirrorExecutor, schema string, tables pgMirrorTables, workbookID, sheet string, cells []pgMirrorCell, batchSize int) error {
	deleteSQL := fmt.Sprintf(`DELETE FROM %s WHERE workbook_id = $1 AND sheet_name = $2`, qualifyPGTable(schema, tables.lookup))
	if _, err := exec.ExecContext(ctx, deleteSQL, workbookID, sheet); err != nil {
		return err
	}
	if len(cells) == 0 {
		return nil
	}
	for start := 0; start < len(cells); start += batchSize {
		end := start + batchSize
		if end > len(cells) {
			end = len(cells)
		}
		if err := upsertPGLookupRows(ctx, exec, schema, tables, workbookID, sheet, cells[start:end]); err != nil {
			return err
		}
	}
	return nil
}

func upsertPGMirroredCells(ctx context.Context, exec pgMirrorExecutor, schema string, tables pgMirrorTables, workbookID, sheet string, cells []pgMirrorCell) error {
	if len(cells) == 0 {
		return nil
	}
	insertSQL, args := buildPGCellUpsertSQL(schema, tables, workbookID, sheet, cells)
	_, err := exec.ExecContext(ctx, insertSQL, args...)
	return err
}

func deletePGMirroredCell(ctx context.Context, exec pgMirrorExecutor, schema string, tables pgMirrorTables, workbookID, sheet, cell string) error {
	sqlStr := fmt.Sprintf(`DELETE FROM %s WHERE workbook_id = $1 AND sheet_name = $2 AND cell_ref = $3`, qualifyPGTable(schema, tables.cell))
	_, err := exec.ExecContext(ctx, sqlStr, workbookID, sheet, cell)
	return err
}

func deletePGMirroredCells(ctx context.Context, exec pgMirrorExecutor, schema string, tables pgMirrorTables, workbookID, sheet string, cells []string) error {
	if len(cells) == 0 {
		return nil
	}
	sqlStr := fmt.Sprintf(`DELETE FROM %s WHERE workbook_id = $1 AND sheet_name = $2 AND cell_ref = ANY($3::text[])`, qualifyPGTable(schema, tables.cell))
	_, err := exec.ExecContext(ctx, sqlStr, workbookID, sheet, cells)
	return err
}

func deletePGLookupRow(ctx context.Context, exec pgMirrorExecutor, schema string, tables pgMirrorTables, workbookID, sheet, cell string) error {
	sqlStr := fmt.Sprintf(`DELETE FROM %s WHERE workbook_id = $1 AND sheet_name = $2 AND cell_ref = $3`, qualifyPGTable(schema, tables.lookup))
	_, err := exec.ExecContext(ctx, sqlStr, workbookID, sheet, cell)
	return err
}

func deletePGLookupRows(ctx context.Context, exec pgMirrorExecutor, schema string, tables pgMirrorTables, workbookID, sheet string, cells []string) error {
	if len(cells) == 0 {
		return nil
	}
	sqlStr := fmt.Sprintf(`DELETE FROM %s WHERE workbook_id = $1 AND sheet_name = $2 AND cell_ref = ANY($3::text[])`, qualifyPGTable(schema, tables.lookup))
	_, err := exec.ExecContext(ctx, sqlStr, workbookID, sheet, cells)
	return err
}

func deletePGRemovedSheets(ctx context.Context, exec pgMirrorExecutor, schema string, tables pgMirrorTables, workbookID string, sheets []string) error {
	sqlStr, args := buildPGDeleteRemovedSheetsSQL(schema, tables, workbookID, sheets)
	_, err := exec.ExecContext(ctx, sqlStr, args...)
	return err
}

func buildPGDeleteRemovedSheetsSQL(schema string, tables pgMirrorTables, workbookID string, sheets []string) (string, []interface{}) {
	sqlStr := fmt.Sprintf(`DELETE FROM %s WHERE workbook_id = $1`, qualifyPGTable(schema, tables.sheet))
	args := []interface{}{workbookID}
	if len(sheets) == 0 {
		return sqlStr, args
	}
	placeholders := make([]string, 0, len(sheets))
	for i, sheet := range sheets {
		placeholders = append(placeholders, fmt.Sprintf("$%d", i+2))
		args = append(args, sheet)
	}
	sqlStr += " AND sheet_name NOT IN (" + strings.Join(placeholders, ",") + ")"
	return sqlStr, args
}

func upsertPGLookupRows(ctx context.Context, exec pgMirrorExecutor, schema string, tables pgMirrorTables, workbookID, sheet string, cells []pgMirrorCell) error {
	if len(cells) == 0 {
		return nil
	}
	insertSQL, args := buildPGLookupUpsertSQL(schema, tables, workbookID, sheet, cells)
	_, err := exec.ExecContext(ctx, insertSQL, args...)
	return err
}

func buildPGLookupUpsertSQL(schema string, tables pgMirrorTables, workbookID, sheet string, cells []pgMirrorCell) (string, []interface{}) {
	cellRefs := make([]string, 0, len(cells))
	rowNums := make([]int32, 0, len(cells))
	colNums := make([]int32, 0, len(cells))
	values := make([]string, 0, len(cells))
	valueTypes := make([]string, 0, len(cells))
	for _, record := range cells {
		cellRefs = append(cellRefs, record.Cell)
		rowNums = append(rowNums, int32(record.Row))
		colNums = append(colNums, int32(record.Col))
		values = append(values, record.Value)
		valueTypes = append(valueTypes, record.ValueType)
	}
	sqlStr := fmt.Sprintf(`
INSERT INTO %s(workbook_id, sheet_name, cell_ref, row_num, col_num, cell_value, cell_value_type, updated_at)
SELECT $1, $2, src.cell_ref, src.row_num, src.col_num, src.cell_value, src.cell_value_type, NOW()
FROM UNNEST($3::text[], $4::int4[], $5::int4[], $6::text[], $7::text[])
	AS src(cell_ref, row_num, col_num, cell_value, cell_value_type)
ON CONFLICT (workbook_id, sheet_name, cell_ref)
DO UPDATE SET
	row_num = EXCLUDED.row_num,
	col_num = EXCLUDED.col_num,
	cell_value = EXCLUDED.cell_value,
	cell_value_type = EXCLUDED.cell_value_type,
	updated_at = NOW()`,
		qualifyPGTable(schema, tables.lookup),
	)
	return sqlStr, []interface{}{workbookID, sheet, cellRefs, rowNums, colNums, values, valueTypes}
}

func buildPGCellUpsertSQL(schema string, tables pgMirrorTables, workbookID, sheet string, cells []pgMirrorCell) (string, []interface{}) {
	cellRefs := make([]string, 0, len(cells))
	rowNums := make([]int32, 0, len(cells))
	colNums := make([]int32, 0, len(cells))
	values := make([]string, 0, len(cells))
	valueTypes := make([]string, 0, len(cells))
	formulas := make([]string, 0, len(cells))
	for _, record := range cells {
		cellRefs = append(cellRefs, record.Cell)
		rowNums = append(rowNums, int32(record.Row))
		colNums = append(colNums, int32(record.Col))
		values = append(values, record.Value)
		valueTypes = append(valueTypes, record.ValueType)
		formulas = append(formulas, record.Formula)
	}
	sqlStr := fmt.Sprintf(`
INSERT INTO %s(workbook_id, sheet_name, cell_ref, row_num, col_num, value, value_type, formula, updated_at)
SELECT $1, $2, src.cell_ref, src.row_num, src.col_num, src.value, src.value_type, src.formula, NOW()
FROM UNNEST($3::text[], $4::int4[], $5::int4[], $6::text[], $7::text[], $8::text[])
	AS src(cell_ref, row_num, col_num, value, value_type, formula)
ON CONFLICT (workbook_id, sheet_name, cell_ref)
DO UPDATE SET
	row_num = EXCLUDED.row_num,
	col_num = EXCLUDED.col_num,
	value = EXCLUDED.value,
	value_type = EXCLUDED.value_type,
	formula = EXCLUDED.formula,
	updated_at = NOW()`,
		qualifyPGTable(schema, tables.cell),
	)
	return sqlStr, []interface{}{workbookID, sheet, cellRefs, rowNums, colNums, values, valueTypes, formulas}
}

func buildPGSingleCellFastPathSQL(schema string, tables pgMirrorTables, workbookID, sheet string, cell pgMirrorCell, includeSheetMeta bool) (string, []interface{}) {
	args := make([]interface{}, 0, 11)
	ctes := make([]string, 0, 3)

	if includeSheetMeta {
		args = append(args, workbookID, sheet, cell.Row, cell.Col)
		ctes = append(ctes, fmt.Sprintf(`sheet_meta AS (
INSERT INTO %s AS tgt(workbook_id, sheet_name, row_count, col_count, cell_count, last_synced_at)
VALUES ($1, $2, $3, $4, -1, NOW())
ON CONFLICT (workbook_id, sheet_name)
DO UPDATE SET
	row_count = GREATEST(tgt.row_count, EXCLUDED.row_count),
	col_count = GREATEST(tgt.col_count, EXCLUDED.col_count),
	cell_count = tgt.cell_count,
	last_synced_at = NOW()
RETURNING 1
)`, qualifyPGTable(schema, tables.sheet)))
	}

	base := len(args) + 1
	if cell.Value == "" && cell.Formula == "" {
		args = append(args, workbookID, sheet, cell.Cell)
		ctes = append(ctes, fmt.Sprintf(`cell_delete AS (
DELETE FROM %s
WHERE workbook_id = $%d AND sheet_name = $%d AND cell_ref = $%d
RETURNING 1
)`, qualifyPGTable(schema, tables.cell), base, base+1, base+2))
		ctes = append(ctes, fmt.Sprintf(`lookup_delete AS (
DELETE FROM %s
WHERE workbook_id = $%d AND sheet_name = $%d AND cell_ref = $%d
RETURNING 1
)`, qualifyPGTable(schema, tables.lookup), base, base+1, base+2))
	} else {
		args = append(args,
			workbookID,
			sheet,
			cell.Cell,
			cell.Row,
			cell.Col,
			cell.Value,
			cell.ValueType,
			cell.Formula,
		)
		ctes = append(ctes, fmt.Sprintf(`cell_upsert AS (
INSERT INTO %s(workbook_id, sheet_name, cell_ref, row_num, col_num, value, value_type, formula, updated_at)
VALUES ($%d, $%d, $%d, $%d, $%d, $%d, $%d, $%d, NOW())
ON CONFLICT (workbook_id, sheet_name, cell_ref)
DO UPDATE SET
	row_num = EXCLUDED.row_num,
	col_num = EXCLUDED.col_num,
	value = EXCLUDED.value,
	value_type = EXCLUDED.value_type,
	formula = EXCLUDED.formula,
	updated_at = NOW()
RETURNING 1
)`, qualifyPGTable(schema, tables.cell), base, base+1, base+2, base+3, base+4, base+5, base+6, base+7))
		ctes = append(ctes, fmt.Sprintf(`lookup_upsert AS (
INSERT INTO %s(workbook_id, sheet_name, cell_ref, row_num, col_num, cell_value, cell_value_type, updated_at)
VALUES ($%d, $%d, $%d, $%d, $%d, $%d, $%d, NOW())
ON CONFLICT (workbook_id, sheet_name, cell_ref)
DO UPDATE SET
	row_num = EXCLUDED.row_num,
	col_num = EXCLUDED.col_num,
	cell_value = EXCLUDED.cell_value,
	cell_value_type = EXCLUDED.cell_value_type,
	updated_at = NOW()
RETURNING 1
)`, qualifyPGTable(schema, tables.lookup), base, base+1, base+2, base+3, base+4, base+5, base+6))
	}

	return "WITH " + strings.Join(ctes, ",\n") + "\nSELECT 1", args
}
