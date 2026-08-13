package excelize

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

// AutoFilterResult holds the complete AutoFilter information read from a worksheet.
type AutoFilterResult struct {
	Ref           string
	FilterColumns []AutoFilterColumnResult
}

// AutoFilterColumnResult holds filter criteria for a single column.
type AutoFilterColumnResult struct {
	ColID         int
	Filters       *FiltersResult
	CustomFilters *CustomFiltersResult
	ColorFilter   *ColorFilterResult
	ColorFilters  *ColorFiltersResult
}

// FiltersResult holds value-list filter criteria.
type FiltersResult struct {
	Blank        bool
	FilterValues []string
}

// CustomFiltersResult holds custom (operator-based) filter criteria.
type CustomFiltersResult struct {
	And   bool
	Items []CustomFilterItemResult
}

// CustomFilterItemResult holds a single custom filter condition.
type CustomFilterItemResult struct {
	Operator string
	Val      string
}

// ColorFilterResult holds OOXML color-based filter criteria.
type ColorFilterResult struct {
	CellColor bool
	DxfID     int
}

// ColorFiltersResult holds Univer-style color filter criteria (direct color values).
type ColorFiltersResult struct {
	CellFillColors []string
}

// rowOnlyRefPattern matches row-only references like "1:5", "1:1048576".
var rowOnlyRefPattern = regexp.MustCompile(`^(\d+):(\d+)$`)

// normalizeAutoFilterRef converts row-only references (e.g., "1:5") to cell
// range references (e.g., "A1:E5") using the sheet's actual dimension to
// determine the last column. Dollar signs are stripped. If the ref is already
// a valid cell range, it is returned as-is.
func (f *File) normalizeAutoFilterRef(sheet, ref string) (string, error) {
	ref = strings.ReplaceAll(ref, "$", "")
	matches := rowOnlyRefPattern.FindStringSubmatch(ref)
	if matches == nil {
		return ref, nil
	}
	startRow, _ := strconv.Atoi(matches[1])
	endRow, _ := strconv.Atoi(matches[2])

	lastCol := "A"
	dim, err := f.GetSheetDimension(sheet)
	if err == nil && dim != "" {
		parts := strings.Split(dim, ":")
		if len(parts) == 2 {
			col := strings.TrimRight(parts[1], "0123456789")
			if col != "" {
				lastCol = col
			}
		}
	}
	return fmt.Sprintf("A%d:%s%d", startRow, lastCol, endRow), nil
}

// GetAutoFilter returns the AutoFilter information for the given worksheet.
// If no AutoFilter is set, it returns (nil, nil).
func (f *File) GetAutoFilter(sheet string) (*AutoFilterResult, error) {
	ws, err := f.workSheetReader(sheet)
	if err != nil {
		return nil, err
	}
	if ws.AutoFilter == nil || ws.AutoFilter.Ref == "" {
		return nil, nil
	}
	result := &AutoFilterResult{
		Ref: strings.ReplaceAll(ws.AutoFilter.Ref, "$", ""),
	}
	for _, fc := range ws.AutoFilter.FilterColumn {
		col := AutoFilterColumnResult{ColID: fc.ColID}
		if fc.Filters != nil {
			col.Filters = &FiltersResult{Blank: fc.Filters.Blank}
			for _, fv := range fc.Filters.Filter {
				col.Filters.FilterValues = append(col.Filters.FilterValues, fv.Val)
			}
		}
		if fc.CustomFilters != nil {
			col.CustomFilters = &CustomFiltersResult{And: fc.CustomFilters.And}
			for _, cf := range fc.CustomFilters.CustomFilter {
				col.CustomFilters.Items = append(col.CustomFilters.Items, CustomFilterItemResult{
					Operator: cf.Operator,
					Val:      cf.Val,
				})
			}
		}
		if fc.ColorFilter != nil {
			col.ColorFilter = &ColorFilterResult{
				CellColor: fc.ColorFilter.CellColor,
				DxfID:     fc.ColorFilter.DxfID,
			}
			// Also populate ColorFilters (Univer-style) by reading the dxf fill color directly
			if fc.ColorFilter.CellColor {
				if color := f.getDxfFillColor(fc.ColorFilter.DxfID); color != "" {
					col.ColorFilters = &ColorFiltersResult{
						CellFillColors: []string{color},
					}
				}
			}
		}
		result.FilterColumns = append(result.FilterColumns, col)
	}
	return result, nil
}

// SetAutoFilterFull sets the AutoFilter on a worksheet with full column filter
// details. It uses the existing AutoFilter method to handle the range and
// DefinedNames, then directly constructs FilterColumn entries on the internal
// worksheet structure — no reflection needed.
func (f *File) SetAutoFilterFull(sheet, ref string, columns []AutoFilterColumnResult) error {
	// Normalize row-only refs like "1:5" to "A1:E5"
	normalizedRef, err := f.normalizeAutoFilterRef(sheet, ref)
	if err != nil {
		return fmt.Errorf("failed to normalize auto filter ref: %w", err)
	}

	// Use existing AutoFilter to set ref and DefinedName
	if err := f.AutoFilter(sheet, normalizedRef, nil); err != nil {
		return fmt.Errorf("failed to set auto filter range: %w", err)
	}

	if len(columns) == 0 {
		return nil
	}

	ws, err := f.workSheetReader(sheet)
	if err != nil {
		return err
	}
	if ws.AutoFilter == nil {
		return fmt.Errorf("AutoFilter unexpectedly nil after setting range")
	}

	filterColumns := make([]*xlsxFilterColumn, 0, len(columns))
	for _, col := range columns {
		fc := &xlsxFilterColumn{ColID: col.ColID}

		if col.Filters != nil {
			fc.Filters = &xlsxFilters{Blank: col.Filters.Blank}
			for _, v := range col.Filters.FilterValues {
				fc.Filters.Filter = append(fc.Filters.Filter, &xlsxFilter{Val: v})
			}
		}

		if col.CustomFilters != nil {
			fc.CustomFilters = &xlsxCustomFilters{And: col.CustomFilters.And}
			for _, item := range col.CustomFilters.Items {
				fc.CustomFilters.CustomFilter = append(fc.CustomFilters.CustomFilter, &xlsxCustomFilter{
					Operator: item.Operator,
					Val:      item.Val,
				})
			}
		}

		if col.ColorFilter != nil {
			fc.ColorFilter = &xlsxColorFilter{
				CellColor: col.ColorFilter.CellColor,
				DxfID:     col.ColorFilter.DxfID,
			}
		} else if col.ColorFilters != nil && len(col.ColorFilters.CellFillColors) > 0 {
			// Convert Univer-style color filter to OOXML dxf entry
			dxfID, err := f.NewConditionalStyle(&Style{
				Fill: Fill{
					Type:    "pattern",
					Pattern: 1, // solid fill
					Color:   []string{col.ColorFilters.CellFillColors[0]},
				},
			})
			if err == nil {
				fc.ColorFilter = &xlsxColorFilter{
					CellColor: true,
					DxfID:     dxfID,
				}
			}
		}

		filterColumns = append(filterColumns, fc)
	}
	ws.AutoFilter.FilterColumn = filterColumns
	return nil
}

// RemoveAutoFilterFull removes the AutoFilter from a worksheet, clearing the
// filter, sort state, filterMode, and the corresponding DefinedName entry.
func (f *File) RemoveAutoFilterFull(sheet string) error {
	ws, err := f.workSheetReader(sheet)
	if err != nil {
		return err
	}

	// Clear AutoFilter and its dependent worksheet state.
	ws.AutoFilter = nil
	ws.SortState = nil

	// Clear filterMode
	if ws.SheetPr != nil {
		ws.SheetPr.FilterMode = false
	}

	// Remove _xlnm._FilterDatabase DefinedName for this sheet
	wb, err := f.workbookReader()
	if err != nil {
		return err
	}
	sheetID, err := f.GetSheetIndex(sheet)
	if err != nil {
		return err
	}
	if wb.DefinedNames != nil {
		filtered := wb.DefinedNames.DefinedName[:0]
		for _, dn := range wb.DefinedNames.DefinedName {
			localID := 0
			if dn.LocalSheetID != nil {
				localID = *dn.LocalSheetID
			}
			if dn.Name == builtInDefinedNames[3] && localID == sheetID {
				continue // skip — this is the filter entry
			}
			filtered = append(filtered, dn)
		}
		wb.DefinedNames.DefinedName = filtered
	}
	return nil
}

// getDxfFillColor reads the fill color directly from the dxf entry at the
// given index, bypassing theme color resolution. Returns "#rrggbb" or "".
func (f *File) getDxfFillColor(dxfID int) string {
	f.mu.Lock()
	s, err := f.stylesReader()
	f.mu.Unlock()
	if err != nil || s.Dxfs == nil || dxfID < 0 || dxfID >= len(s.Dxfs.Dxfs) {
		return ""
	}
	dxf := s.Dxfs.Dxfs[dxfID]
	if dxf.Fill == nil || dxf.Fill.PatternFill == nil {
		return ""
	}
	pf := dxf.Fill.PatternFill
	var rgb string
	if pf.FgColor != nil && pf.FgColor.RGB != "" {
		rgb = pf.FgColor.RGB
	} else if pf.BgColor != nil && pf.BgColor.RGB != "" {
		rgb = pf.BgColor.RGB
	}
	if rgb == "" {
		return ""
	}
	// Strip the "FF" alpha prefix (AARRGGBB → RRGGBB)
	rgb = strings.ToUpper(rgb)
	if len(rgb) == 8 {
		rgb = strings.TrimPrefix(rgb, "FF")
	}
	if len(rgb) == 6 {
		return "#" + strings.ToLower(rgb)
	}
	return ""
}
