package excelize

import (
	"encoding/xml"
	"fmt"
	"path/filepath"
	"strings"
	"testing"

	_ "image/jpeg"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

func TestAdjustMergeCells(t *testing.T) {
	f := NewFile()
	// Test adjustAutoFilter with illegal cell reference
	assert.Equal(t, f.adjustMergeCells(&xlsxWorksheet{
		MergeCells: &xlsxMergeCells{
			Cells: []*xlsxMergeCell{
				{
					Ref: "A:B1",
				},
			},
		},
	}, "Sheet1", rows, 0, 0, 1), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")))
	assert.Equal(t, f.adjustMergeCells(&xlsxWorksheet{
		MergeCells: &xlsxMergeCells{
			Cells: []*xlsxMergeCell{
				{
					Ref: "A1:B",
				},
			},
		},
	}, "Sheet1", rows, 0, 0, 1), newCellNameToCoordinatesError("B", newInvalidCellNameError("B")))
	assert.NoError(t, f.adjustMergeCells(&xlsxWorksheet{
		MergeCells: &xlsxMergeCells{
			Cells: []*xlsxMergeCell{
				{
					Ref: "A1:B1",
				},
			},
		},
	}, "Sheet1", rows, 1, -1, 1))
	assert.NoError(t, f.adjustMergeCells(&xlsxWorksheet{
		MergeCells: &xlsxMergeCells{
			Cells: []*xlsxMergeCell{
				{
					Ref: "A1:A2",
				},
			},
		},
	}, "Sheet1", columns, 1, -1, 1))
	assert.NoError(t, f.adjustMergeCells(&xlsxWorksheet{
		MergeCells: &xlsxMergeCells{
			Cells: []*xlsxMergeCell{
				{
					Ref: "A2",
				},
			},
		},
	}, "Sheet1", columns, 1, -1, 1))

	// Test adjust merge cells
	var cases []struct {
		label      string
		ws         *xlsxWorksheet
		dir        adjustDirection
		num        int
		offset     int
		expect     string
		expectRect []int
	}

	// Test adjust merged cell when insert rows and columns
	cases = []struct {
		label      string
		ws         *xlsxWorksheet
		dir        adjustDirection
		num        int
		offset     int
		expect     string
		expectRect []int
	}{
		{
			label: "insert row on ref",
			ws: &xlsxWorksheet{
				MergeCells: &xlsxMergeCells{
					Cells: []*xlsxMergeCell{
						{
							Ref:  "A2:B3",
							rect: []int{1, 2, 2, 3},
						},
					},
				},
			},
			dir:        rows,
			num:        2,
			offset:     1,
			expect:     "A3:B4",
			expectRect: []int{1, 3, 2, 4},
		},
		{
			label: "insert row on bottom of ref",
			ws: &xlsxWorksheet{
				MergeCells: &xlsxMergeCells{
					Cells: []*xlsxMergeCell{
						{
							Ref:  "A2:B3",
							rect: []int{1, 2, 2, 3},
						},
					},
				},
			},
			dir:        rows,
			num:        3,
			offset:     1,
			expect:     "A2:B4",
			expectRect: []int{1, 2, 2, 4},
		},
		{
			label: "insert column on the left",
			ws: &xlsxWorksheet{
				MergeCells: &xlsxMergeCells{
					Cells: []*xlsxMergeCell{
						{
							Ref:  "A2:B3",
							rect: []int{1, 2, 2, 3},
						},
					},
				},
			},
			dir:        columns,
			num:        1,
			offset:     1,
			expect:     "B2:C3",
			expectRect: []int{2, 2, 3, 3},
		},
	}
	for _, c := range cases {
		assert.NoError(t, f.adjustMergeCells(c.ws, "Sheet1", c.dir, c.num, 1, 1))
		assert.Equal(t, c.expect, c.ws.MergeCells.Cells[0].Ref, c.label)
		assert.Equal(t, c.expectRect, c.ws.MergeCells.Cells[0].rect, c.label)
	}

	// Test adjust merged cells when delete rows and columns
	cases = []struct {
		label      string
		ws         *xlsxWorksheet
		dir        adjustDirection
		num        int
		offset     int
		expect     string
		expectRect []int
	}{
		{
			label: "delete row on top of ref",
			ws: &xlsxWorksheet{
				MergeCells: &xlsxMergeCells{
					Cells: []*xlsxMergeCell{
						{
							Ref:  "A2:B3",
							rect: []int{1, 2, 2, 3},
						},
					},
				},
			},
			dir:        rows,
			num:        2,
			offset:     -1,
			expect:     "A2:B2",
			expectRect: []int{1, 2, 2, 2},
		},
		{
			label: "delete row on bottom of ref",
			ws: &xlsxWorksheet{
				MergeCells: &xlsxMergeCells{
					Cells: []*xlsxMergeCell{
						{
							Ref:  "A2:B3",
							rect: []int{1, 2, 2, 3},
						},
					},
				},
			},
			dir:        rows,
			num:        3,
			offset:     -1,
			expect:     "A2:B2",
			expectRect: []int{1, 2, 2, 2},
		},
		{
			label: "delete column on the ref left",
			ws: &xlsxWorksheet{
				MergeCells: &xlsxMergeCells{
					Cells: []*xlsxMergeCell{
						{
							Ref:  "A2:B3",
							rect: []int{1, 2, 2, 3},
						},
					},
				},
			},
			dir:        columns,
			num:        1,
			offset:     -1,
			expect:     "A2:A3",
			expectRect: []int{1, 2, 1, 3},
		},
		{
			label: "delete column on the ref right",
			ws: &xlsxWorksheet{
				MergeCells: &xlsxMergeCells{
					Cells: []*xlsxMergeCell{
						{
							Ref:  "A2:B3",
							rect: []int{1, 2, 2, 3},
						},
					},
				},
			},
			dir:        columns,
			num:        2,
			offset:     -1,
			expect:     "A2:A3",
			expectRect: []int{1, 2, 1, 3},
		},
	}
	for _, c := range cases {
		assert.NoError(t, f.adjustMergeCells(c.ws, "Sheet1", c.dir, c.num, -1, 1))
		assert.Equal(t, c.expect, c.ws.MergeCells.Cells[0].Ref, c.label)
	}

	// Test delete one row or column
	cases = []struct {
		label      string
		ws         *xlsxWorksheet
		dir        adjustDirection
		num        int
		offset     int
		expect     string
		expectRect []int
	}{
		{
			label: "delete one row ref",
			ws: &xlsxWorksheet{
				MergeCells: &xlsxMergeCells{
					Cells: []*xlsxMergeCell{
						{
							Ref:  "A1:B1",
							rect: []int{1, 1, 2, 1},
						},
					},
				},
			},
			dir:    rows,
			num:    1,
			offset: -1,
		},
		{
			label: "delete one column ref",
			ws: &xlsxWorksheet{
				MergeCells: &xlsxMergeCells{
					Cells: []*xlsxMergeCell{
						{
							Ref:  "A1:A2",
							rect: []int{1, 1, 1, 2},
						},
					},
				},
			},
			dir:    columns,
			num:    1,
			offset: -1,
		},
	}
	for _, c := range cases {
		assert.NoError(t, f.adjustMergeCells(c.ws, "Sheet1", c.dir, c.num, -1, 1))
		assert.Len(t, c.ws.MergeCells.Cells, 0, c.label)
	}

	f = NewFile()
	p1, p2 := f.adjustMergeCellsHelper(2, 1, 0, 0)
	assert.Equal(t, 1, p1)
	assert.Equal(t, 2, p2)
	f.deleteMergeCell(nil, -1)
}

func TestAdjustAutoFilter(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.adjustAutoFilter(&xlsxWorksheet{
		SheetData: xlsxSheetData{
			Row: []xlsxRow{{Hidden: true, R: 2}},
		},
		AutoFilter: &xlsxAutoFilter{
			Ref: "A1:A3",
		},
	}, "Sheet1", rows, 1, -1, 1))
	// Malformed legacy filters must not block unrelated structural operations.
	worksheet, err := f.workSheetReader("Sheet1")
	require.NoError(t, err)
	worksheet.AutoFilter = &xlsxAutoFilter{Ref: "A:B1"}
	assert.NoError(t, f.adjustAutoFilter(worksheet, "Sheet1", rows, 0, 0, 1))
	assert.Nil(t, worksheet.AutoFilter)
	worksheet.AutoFilter = &xlsxAutoFilter{Ref: "A1:B"}
	assert.NoError(t, f.adjustAutoFilter(worksheet, "Sheet1", rows, 0, 0, 1))
	assert.Nil(t, worksheet.AutoFilter)

	t.Run("preserves advanced criteria", func(t *testing.T) {
		worksheet := &xlsxWorksheet{AutoFilter: &xlsxAutoFilter{
			Ref: "E1:G4",
			FilterColumn: []*xlsxFilterColumn{{
				ColID:         1,
				HiddenButton:  true,
				ShowButton:    true,
				DynamicFilter: &xlsxDynamicFilter{Type: "aboveAverage"},
				IconFilter:    &xlsxIconFilter{IconID: 2, IconSet: "3Arrows"},
				Top10:         &xlsxTop10{Top: true, Val: 10},
				Filters: &xlsxFilters{
					CalendarType:  "gregorian",
					DateGroupItem: []*xlsxDateGroupItem{{Year: 2026, Month: 8, DateTimeGrouping: "month"}},
				},
			}},
		}}
		assert.NoError(t, f.adjustAutoFilter(worksheet, "Sheet1", columns, 2, -2, 1))
		assert.Equal(t, "C1:E4", worksheet.AutoFilter.Ref)
		column := worksheet.AutoFilter.FilterColumn[0]
		assert.Equal(t, 1, column.ColID)
		assert.True(t, column.HiddenButton)
		assert.True(t, column.ShowButton)
		assert.Equal(t, "aboveAverage", column.DynamicFilter.Type)
		assert.Equal(t, "3Arrows", column.IconFilter.IconSet)
		assert.Equal(t, float64(10), column.Top10.Val)
		assert.Equal(t, "gregorian", column.Filters.CalendarType)
	})

	t.Run("removes filter metadata when header row is deleted", func(t *testing.T) {
		file := NewFile()
		require.NoError(t, file.AutoFilter("Sheet1", "A1:A3", nil))
		worksheet, err := file.workSheetReader("Sheet1")
		require.NoError(t, err)
		worksheet.SheetData.Row = []xlsxRow{{R: 1}, {R: 2, Hidden: true}, {R: 3}}
		require.NoError(t, file.RemoveRows("Sheet1", 1, 1))
		assert.Nil(t, worksheet.AutoFilter)
		assert.True(t, worksheet.SheetData.Row[0].Hidden)
		workbook, err := file.workbookReader()
		require.NoError(t, err)
		for _, name := range workbook.DefinedNames.DefinedName {
			assert.NotEqual(t, builtInDefinedNames[3], name.Name)
		}
	})

	t.Run("keeps filter when its first column is deleted", func(t *testing.T) {
		worksheet := &xlsxWorksheet{AutoFilter: &xlsxAutoFilter{Ref: "E1:G4", FilterColumn: []*xlsxFilterColumn{{ColID: 1}}}}
		assert.NoError(t, f.adjustAutoFilter(worksheet, "Sheet1", columns, 5, -1, 1))
		assert.Equal(t, "E1:F4", worksheet.AutoFilter.Ref)
		assert.Equal(t, 0, worksheet.AutoFilter.FilterColumn[0].ColID)
	})

	t.Run("updates sort state with filter", func(t *testing.T) {
		worksheet := &xlsxWorksheet{AutoFilter: &xlsxAutoFilter{Ref: "A1:C4"}, SortState: &xlsxSortState{Ref: "A1:C4"}}
		assert.NoError(t, f.adjustAutoFilter(worksheet, "Sheet1", rows, 2, 2, 1))
		assert.Equal(t, "A1:C6", worksheet.AutoFilter.Ref)
		assert.Equal(t, "A1:C6", worksheet.SortState.Ref)
	})
}

func TestAdjustAutoFilterStructuralOperations(t *testing.T) {
	newFilteredFile := func(t *testing.T) *File {
		t.Helper()
		file := NewFile()
		for row := 1; row <= 4; row++ {
			for col := 5; col <= 7; col++ {
				cell, err := CoordinatesToCellName(col, row)
				require.NoError(t, err)
				require.NoError(t, file.SetCellValue("Sheet1", cell, cell))
			}
		}
		require.NoError(t, file.AutoFilter("Sheet1", "E1:G4", nil))
		worksheet, err := file.workSheetReader("Sheet1")
		require.NoError(t, err)
		worksheet.AutoFilter.FilterColumn = []*xlsxFilterColumn{{
			ColID:         1,
			HiddenButton:  true,
			ShowButton:    true,
			DynamicFilter: &xlsxDynamicFilter{Type: "aboveAverage"},
			IconFilter:    &xlsxIconFilter{IconID: 2, IconSet: "3Arrows"},
			Top10:         &xlsxTop10{Top: true, Val: 10},
			Filters: &xlsxFilters{
				CalendarType:  "gregorian",
				DateGroupItem: []*xlsxDateGroupItem{{Year: 2026, Month: 8, DateTimeGrouping: "month"}},
			},
		}}
		worksheet.SortState = &xlsxSortState{Ref: "E1:G4"}
		return file
	}

	assertSavedFilter := func(t *testing.T, file *File, ref, sortRef string, colID int) {
		t.Helper()
		path := filepath.Join(t.TempDir(), "autofilter.xlsx")
		require.NoError(t, file.SaveAs(path))
		reopened, err := OpenFile(path)
		require.NoError(t, err)
		t.Cleanup(func() { _ = reopened.Close() })
		worksheet, err := reopened.workSheetReader("Sheet1")
		require.NoError(t, err)
		require.NotNil(t, worksheet.AutoFilter)
		assert.Equal(t, ref, worksheet.AutoFilter.Ref)
		require.Len(t, worksheet.AutoFilter.FilterColumn, 1)
		column := worksheet.AutoFilter.FilterColumn[0]
		assert.Equal(t, colID, column.ColID)
		assert.True(t, column.HiddenButton)
		assert.True(t, column.ShowButton)
		assert.Equal(t, "aboveAverage", column.DynamicFilter.Type)
		assert.Equal(t, "3Arrows", column.IconFilter.IconSet)
		assert.Equal(t, float64(10), column.Top10.Val)
		assert.Equal(t, "gregorian", column.Filters.CalendarType)
		require.NotNil(t, worksheet.SortState)
		assert.Equal(t, sortRef, worksheet.SortState.Ref)
	}

	t.Run("insert rows", func(t *testing.T) {
		file := newFilteredFile(t)
		t.Cleanup(func() { _ = file.Close() })
		require.NoError(t, file.InsertRows("Sheet1", 2, 2))
		assertSavedFilter(t, file, "E1:G6", "E1:G6", 1)
	})

	t.Run("remove rows including header", func(t *testing.T) {
		file := newFilteredFile(t)
		t.Cleanup(func() { _ = file.Close() })
		worksheet, err := file.workSheetReader("Sheet1")
		require.NoError(t, err)
		worksheet.SheetData.Row[1].Hidden = true
		require.NoError(t, file.RemoveRows("Sheet1", 1, 1))
		path := filepath.Join(t.TempDir(), "autofilter-removed.xlsx")
		require.NoError(t, file.SaveAs(path))
		reopened, err := OpenFile(path)
		require.NoError(t, err)
		t.Cleanup(func() { _ = reopened.Close() })
		worksheet, err = reopened.workSheetReader("Sheet1")
		require.NoError(t, err)
		assert.Nil(t, worksheet.AutoFilter)
		assert.Nil(t, worksheet.SortState)
		assert.True(t, worksheet.SheetData.Row[0].Hidden)
		workbook, err := reopened.workbookReader()
		require.NoError(t, err)
		for _, name := range workbook.DefinedNames.DefinedName {
			assert.NotEqual(t, builtInDefinedNames[3], name.Name)
		}
	})

	t.Run("remove only filter column clears sort state and preserves hidden rows", func(t *testing.T) {
		file := NewFile()
		t.Cleanup(func() { _ = file.Close() })
		for row := 1; row <= 3; row++ {
			require.NoError(t, file.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), row))
			require.NoError(t, file.SetCellValue("Sheet1", fmt.Sprintf("B%d", row), row))
		}
		require.NoError(t, file.AutoFilter("Sheet1", "A1:A3", nil))
		worksheet, err := file.workSheetReader("Sheet1")
		require.NoError(t, err)
		worksheet.SortState = &xlsxSortState{Ref: "A1:A3"}
		worksheet.SheetData.Row[1].Hidden = true
		require.NoError(t, file.RemoveCols("Sheet1", "A", 1))
		path := filepath.Join(t.TempDir(), "autofilter-column-removed.xlsx")
		require.NoError(t, file.SaveAs(path))
		reopened, err := OpenFile(path)
		require.NoError(t, err)
		t.Cleanup(func() { _ = reopened.Close() })
		worksheet, err = reopened.workSheetReader("Sheet1")
		require.NoError(t, err)
		assert.Nil(t, worksheet.AutoFilter)
		assert.Nil(t, worksheet.SortState)
		assert.True(t, worksheet.SheetData.Row[1].Hidden)
	})

	t.Run("insert columns", func(t *testing.T) {
		file := newFilteredFile(t)
		t.Cleanup(func() { _ = file.Close() })
		require.NoError(t, file.InsertCols("Sheet1", "F", 2))
		assertSavedFilter(t, file, "E1:I4", "E1:I4", 3)
	})

	t.Run("remove leftmost column", func(t *testing.T) {
		file := newFilteredFile(t)
		t.Cleanup(func() { _ = file.Close() })
		require.NoError(t, file.RemoveCols("Sheet1", "E", 1))
		assertSavedFilter(t, file, "E1:F4", "E1:F4", 0)
	})

	t.Run("remove rightmost column", func(t *testing.T) {
		file := newFilteredFile(t)
		t.Cleanup(func() { _ = file.Close() })
		require.NoError(t, file.RemoveCols("Sheet1", "G", 1))
		assertSavedFilter(t, file, "E1:F4", "E1:F4", 1)
	})

	t.Run("remove final data row", func(t *testing.T) {
		file := newFilteredFile(t)
		t.Cleanup(func() { _ = file.Close() })
		require.NoError(t, file.RemoveRows("Sheet1", 4, 1))
		assertSavedFilter(t, file, "E1:G3", "E1:G3", 1)
	})

	t.Run("malformed filter does not block insert", func(t *testing.T) {
		file := NewFile()
		t.Cleanup(func() { _ = file.Close() })
		worksheet, err := file.workSheetReader("Sheet1")
		require.NoError(t, err)
		worksheet.AutoFilter = &xlsxAutoFilter{Ref: "A:B1"}
		worksheet.SortState = &xlsxSortState{Ref: "A1:B3"}
		require.NoError(t, file.InsertRows("Sheet1", 2, 1))
		path := filepath.Join(t.TempDir(), "malformed-filter.xlsx")
		require.NoError(t, file.SaveAs(path))
		reopened, err := OpenFile(path)
		require.NoError(t, err)
		t.Cleanup(func() { _ = reopened.Close() })
		result, err := reopened.GetAutoFilter("Sheet1")
		require.NoError(t, err)
		assert.Nil(t, result)
		worksheet, err = reopened.workSheetReader("Sheet1")
		require.NoError(t, err)
		assert.Nil(t, worksheet.SortState)
	})
}

func TestAdjustTable(t *testing.T) {
	f, sheetName := NewFile(), "Sheet1"
	for idx, reference := range []string{"B2:C3", "E3:F5", "H5:H8", "J5:K9"} {
		assert.NoError(t, f.AddTable(sheetName, &Table{
			Range:             reference,
			Name:              fmt.Sprintf("table%d", idx),
			StyleName:         "TableStyleMedium2",
			ShowFirstColumn:   true,
			ShowLastColumn:    true,
			ShowRowStripes:    boolPtr(false),
			ShowColumnStripes: true,
		}))
	}
	assert.NoError(t, f.RemoveRow(sheetName, 2))
	assert.NoError(t, f.RemoveRow(sheetName, 3))
	assert.NoError(t, f.RemoveRow(sheetName, 3))
	assert.NoError(t, f.RemoveCol(sheetName, "H"))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAdjustTable.xlsx")))

	f = NewFile()
	assert.NoError(t, f.AddTable(sheetName, &Table{Range: "A1:D5"}))
	// Test adjust table with non-table part
	f.Pkg.Delete("xl/tables/table1.xml")
	assert.NoError(t, f.RemoveRow(sheetName, 1))
	// Test adjust table with unsupported charset
	f.Pkg.Store("xl/tables/table1.xml", MacintoshCyrillicCharset)
	assert.EqualError(t, f.RemoveRow(sheetName, 1), "XML syntax error on line 1: invalid UTF-8")
	// Test adjust table with invalid table range reference
	f.Pkg.Store("xl/tables/table1.xml", []byte(`<table ref="-" />`))
	assert.Equal(t, ErrParameterInvalid, f.RemoveRow(sheetName, 1))
}

func TestAdjustHelper(t *testing.T) {
	f := NewFile()
	_, err := f.NewSheet("Sheet2")
	assert.NoError(t, err)
	f.Sheet.Store("xl/worksheets/sheet1.xml", &xlsxWorksheet{
		MergeCells: &xlsxMergeCells{Cells: []*xlsxMergeCell{{Ref: "A:B1"}}},
	})
	f.Sheet.Store("xl/worksheets/sheet2.xml", &xlsxWorksheet{
		AutoFilter: &xlsxAutoFilter{Ref: "A1:B"},
	})
	// Invalid merge references still fail, while legacy AutoFilter references
	// are cleared so structural edits can proceed.
	assert.Equal(t, f.adjustHelper("Sheet1", rows, 0, 0), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")))
	assert.NoError(t, f.adjustHelper("Sheet2", rows, 0, 0))
	worksheet, err := f.workSheetReader("Sheet2")
	require.NoError(t, err)
	assert.Nil(t, worksheet.AutoFilter)
	// Test adjustHelper on not exists worksheet
	assert.EqualError(t, f.adjustHelper("SheetN", rows, 0, 0), "sheet SheetN does not exist")
}

func TestAdjustCalcChain(t *testing.T) {
	f := NewFile()
	f.CalcChain = &xlsxCalcChain{
		C: []xlsxCalcChainC{{R: "B2", I: 2}, {R: "B2", I: 1}, {R: "A1", I: 1}},
	}
	assert.NoError(t, f.InsertCols("Sheet1", "A", 1))
	assert.NoError(t, f.InsertRows("Sheet1", 1, 1))

	f.CalcChain = &xlsxCalcChain{
		C: []xlsxCalcChainC{{R: "B2", I: 1}, {R: "B3"}, {R: "A1"}},
	}
	assert.NoError(t, f.RemoveRow("Sheet1", 3))
	assert.NoError(t, f.RemoveCol("Sheet1", "B"))

	f.CalcChain = &xlsxCalcChain{C: []xlsxCalcChainC{{R: "B2", I: 2}, {R: "B2", I: 1}}}
	f.CalcChain.C[1].R = "invalid coordinates"
	assert.Equal(t, f.InsertCols("Sheet1", "A", 1), newCellNameToCoordinatesError("invalid coordinates", newInvalidCellNameError("invalid coordinates")))
	f.CalcChain = nil
	assert.NoError(t, f.InsertCols("Sheet1", "A", 1))
}

func TestAdjustCols(t *testing.T) {
	sheetName := "Sheet1"
	preset := func() (*File, error) {
		f := NewFile()
		if err := f.SetColWidth(sheetName, "J", "T", 5); err != nil {
			return f, err
		}
		if err := f.SetSheetRow(sheetName, "J1", &[]string{"J1", "K1", "L1", "M1", "N1", "O1", "P1", "Q1", "R1", "S1", "T1"}); err != nil {
			return f, err
		}
		return f, nil
	}
	baseTbl := []string{"B", "J", "O", "O", "O", "U", "V"}
	insertTbl := []int{2, 2, 2, 5, 6, 2, 2}
	expectedTbl := []map[string]float64{
		{"J": defaultColWidth, "K": defaultColWidth, "U": 5, "V": 5, "W": defaultColWidth},
		{"J": defaultColWidth, "K": defaultColWidth, "U": 5, "V": 5, "W": defaultColWidth},
		{"O": 5, "P": 5, "U": 5, "V": 5, "W": defaultColWidth},
		{"O": 5, "S": 5, "X": 5, "Y": 5, "Z": defaultColWidth},
		{"O": 5, "S": 5, "Y": 5, "X": 5, "AA": defaultColWidth},
		{"U": 5, "V": 5, "W": defaultColWidth},
		{"U": defaultColWidth, "V": defaultColWidth, "W": defaultColWidth},
	}
	for idx, columnName := range baseTbl {
		f, err := preset()
		assert.NoError(t, err)
		assert.NoError(t, f.InsertCols(sheetName, columnName, insertTbl[idx]))
		for column, expected := range expectedTbl[idx] {
			width, err := f.GetColWidth(sheetName, column)
			assert.NoError(t, err)
			assert.Equal(t, expected, width, column)
		}
		assert.NoError(t, f.Close())
	}

	baseTbl = []string{"B", "J", "O", "T"}
	expectedTbl = []map[string]float64{
		{"H": defaultColWidth, "I": 5, "S": 5, "T": defaultColWidth},
		{"I": defaultColWidth, "J": 5, "S": 5, "T": defaultColWidth},
		{"I": defaultColWidth, "O": 5, "S": 5, "T": defaultColWidth},
		{"R": 5, "S": 5, "T": defaultColWidth, "U": defaultColWidth},
	}
	for idx, columnName := range baseTbl {
		f, err := preset()
		assert.NoError(t, err)
		assert.NoError(t, f.RemoveCol(sheetName, columnName))
		for column, expected := range expectedTbl[idx] {
			width, err := f.GetColWidth(sheetName, column)
			assert.NoError(t, err)
			assert.Equal(t, expected, width, column)
		}
		assert.NoError(t, f.Close())
	}

	f, err := preset()
	assert.NoError(t, err)
	assert.NoError(t, f.SetColWidth(sheetName, "I", "I", 8))
	for i := 0; i <= 12; i++ {
		assert.NoError(t, f.RemoveCol(sheetName, "I"))
	}
	for c := 9; c <= 21; c++ {
		columnName, err := ColumnNumberToName(c)
		assert.NoError(t, err)
		width, err := f.GetColWidth(sheetName, columnName)
		assert.NoError(t, err)
		assert.Equal(t, defaultColWidth, width, columnName)
	}

	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).Cols = nil
	assert.NoError(t, f.RemoveCol(sheetName, "A"))

	assert.NoError(t, f.Close())

	f = NewFile()
	assert.NoError(t, f.SetColWidth("Sheet1", "XFB", "XFC", 12))
	assert.NoError(t, f.InsertCols("Sheet1", "A", 2))
	ws, ok = f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	assert.Equal(t, MaxColumns, ws.(*xlsxWorksheet).Cols.Col[0].Min)
	assert.Equal(t, MaxColumns, ws.(*xlsxWorksheet).Cols.Col[0].Max)

	assert.NoError(t, f.InsertCols("Sheet1", "A", 2))
	assert.Nil(t, ws.(*xlsxWorksheet).Cols)

	f = NewFile()
	assert.NoError(t, f.SetCellFormula("Sheet1", "A2", "(1-0.5)/2"))
	assert.NoError(t, f.InsertCols("Sheet1", "A", 1))
	formula, err := f.GetCellFormula("Sheet1", "B2")
	assert.NoError(t, err)
	assert.Equal(t, "(1-0.5)/2", formula)
	assert.NoError(t, f.Close())
}

func TestAdjustColDimensions(t *testing.T) {
	f := NewFile()
	ws, err := f.workSheetReader("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellFormula("Sheet1", "C3", "A1+B1"))
	assert.Equal(t, ErrColumnNumber, f.adjustColDimensions("Sheet1", ws, 1, MaxColumns))

	// Empty cells at the boundary should be removed instead of returning error
	f2 := NewFile()
	ws2, err := f2.workSheetReader("Sheet1")
	assert.NoError(t, err)
	// Place a non-empty cell within range and an empty cell near MaxColumns
	assert.NoError(t, f2.SetCellValue("Sheet1", "A1", 100))
	ws2.prepareSheetXML(MaxColumns, 1)
	// The cell at MaxColumns is empty (created by prepareSheetXML), inserting
	// a column should silently remove it rather than returning ErrColumnNumber
	assert.NoError(t, f2.adjustColDimensions("Sheet1", ws2, 1, 1))
	// Verify the non-empty cell is still present (shifted from A1 to B1 by offset=1)
	val, err := f2.GetCellValue("Sheet1", "B1")
	assert.NoError(t, err)
	assert.Equal(t, "100", val)
	// Verify empty cell at MaxColumns was removed
	for _, c := range ws2.SheetData.Row[0].C {
		col, _, _ := CellNameToCoordinates(c.R)
		assert.LessOrEqual(t, col, MaxColumns, "no cell should exceed MaxColumns")
	}

	// A non-empty cell at the boundary should still return ErrColumnNumber
	f3 := NewFile()
	ws3, err := f3.workSheetReader("Sheet1")
	assert.NoError(t, err)
	cellName, _ := CoordinatesToCellName(MaxColumns, 1)
	assert.NoError(t, f3.SetCellValue("Sheet1", cellName, "data"))
	assert.Equal(t, ErrColumnNumber, f3.adjustColDimensions("Sheet1", ws3, 1, 1))

	// A cell with formula at the boundary should still return ErrColumnNumber
	f4 := NewFile()
	ws4, err := f4.workSheetReader("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, f4.SetCellFormula("Sheet1", cellName, "1+1"))
	assert.Equal(t, ErrColumnNumber, f4.adjustColDimensions("Sheet1", ws4, 1, 1))

	_, err = f.NewSheet("Sheet2")
	assert.NoError(t, err)
	f.Sheet.Delete("xl/worksheets/sheet2.xml")
	f.Pkg.Store("xl/worksheets/sheet2.xml", MacintoshCyrillicCharset)
	assert.EqualError(t, f.adjustColDimensions("Sheet2", ws, 2, 1), "XML syntax error on line 1: invalid UTF-8")
}

func TestAdjustRowDimensions(t *testing.T) {
	f := NewFile()
	ws, err := f.workSheetReader("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellFormula("Sheet1", "C3", "A1+B1"))
	assert.Equal(t, ErrMaxRows, f.adjustRowDimensions("Sheet1", ws, 1, TotalRows))

	_, err = f.NewSheet("Sheet2")
	assert.NoError(t, err)
	f.Sheet.Delete("xl/worksheets/sheet2.xml")
	f.Pkg.Store("xl/worksheets/sheet2.xml", MacintoshCyrillicCharset)
	assert.EqualError(t, f.adjustRowDimensions("Sheet1", ws, 2, 1), "XML syntax error on line 1: invalid UTF-8")

	f = NewFile()
	_, err = f.NewSheet("Sheet2")
	assert.NoError(t, err)
	ws, err = f.workSheetReader("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellFormula("Sheet1", "B2", fmt.Sprintf("Sheet2!A%d", TotalRows)))
	assert.Equal(t, ErrMaxRows, f.adjustRowDimensions("Sheet2", ws, 1, TotalRows))
}

func TestAdjustHyperlinks(t *testing.T) {
	f := NewFile()
	ws, err := f.workSheetReader("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellFormula("Sheet1", "C3", "A1+B1"))
	f.adjustHyperlinks(ws, "Sheet1", rows, 3, -1)

	// Test adjust hyperlinks location with positive offset
	assert.NoError(t, f.SetCellHyperLink("Sheet1", "F5", "Sheet1!A1", "Location"))
	assert.NoError(t, f.InsertRows("Sheet1", 1, 1))
	link, target, err := f.GetCellHyperLink("Sheet1", "F6")
	assert.NoError(t, err)
	assert.True(t, link)
	assert.Equal(t, target, "Sheet1!A1")

	// Test adjust hyperlinks location with negative offset
	assert.NoError(t, f.RemoveRow("Sheet1", 1))
	link, target, err = f.GetCellHyperLink("Sheet1", "F5")
	assert.NoError(t, err)
	assert.True(t, link)
	assert.Equal(t, target, "Sheet1!A1")

	// Test adjust hyperlinks location on remove row
	assert.NoError(t, f.RemoveRow("Sheet1", 5))
	link, target, err = f.GetCellHyperLink("Sheet1", "F5")
	assert.NoError(t, err)
	assert.False(t, link)
	assert.Empty(t, target)

	// Test adjust hyperlinks location on remove column
	assert.NoError(t, f.SetCellHyperLink("Sheet1", "F5", "Sheet1!A1", "Location"))
	assert.NoError(t, f.RemoveCol("Sheet1", "F"))
	link, target, err = f.GetCellHyperLink("Sheet1", "F5")
	assert.NoError(t, err)
	assert.False(t, link)
	assert.Empty(t, target)

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAdjustHyperlinks.xlsx")))
	assert.NoError(t, f.Close())
}

func TestAdjustFormula(t *testing.T) {
	f := NewFile()
	formulaType, ref := STCellFormulaTypeShared, "C1:C5"
	assert.NoError(t, f.SetCellFormula("Sheet1", "C1", "A1+B1", FormulaOpts{Ref: &ref, Type: &formulaType}))
	assert.NoError(t, f.DuplicateRowTo("Sheet1", 1, 10))
	assert.NoError(t, f.InsertCols("Sheet1", "B", 1))
	assert.NoError(t, f.InsertRows("Sheet1", 1, 1))
	for cell, expected := range map[string]string{"D2": "A2+C2", "D3": "A3+C3", "D11": "A11+C11"} {
		formula, err := f.GetCellFormula("Sheet1", cell)
		assert.NoError(t, err)
		assert.Equal(t, expected, formula)
	}
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAdjustFormula.xlsx")))
	assert.NoError(t, f.Close())

	assert.NoError(t, f.adjustFormula("Sheet1", "Sheet1", &xlsxC{}, rows, 0, 0, false))
	assert.Equal(t, newCellNameToCoordinatesError("-", newInvalidCellNameError("-")), f.adjustFormula("Sheet1", "Sheet1", &xlsxC{F: &xlsxF{Ref: "-"}}, rows, 0, 0, false))
	assert.Equal(t, ErrColumnNumber, f.adjustFormula("Sheet1", "Sheet1", &xlsxC{F: &xlsxF{Ref: "XFD1:XFD1"}}, columns, 0, 1, false))

	// When a formula column reference overflows MaxColumns after adjustment,
	// it should produce #REF! instead of an error (matching Excel behavior)
	result, err := f.adjustFormulaRef("Sheet1", "Sheet1", "XFE1", false, columns, 0, 1, false)
	assert.NoError(t, err)
	assert.Contains(t, result, "#REF!")
	result, err = f.adjustFormulaRef("Sheet1", "Sheet1", "XFD1", false, columns, 0, 1, false)
	assert.NoError(t, err)
	assert.Contains(t, result, "#REF!")

	f = NewFile()
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "XFD1"))
	assert.NoError(t, f.InsertCols("Sheet1", "A", 1))
	formulaResult, _ := f.GetCellFormula("Sheet1", "C1")
	assert.Contains(t, formulaResult, "#REF!")

	assert.NoError(t, f.SetCellFormula("Sheet1", "C2", fmt.Sprintf("A%d", TotalRows)))
	assert.Equal(t, ErrMaxRows, f.InsertRows("Sheet1", 1, 1))

	f = NewFile()
	assert.NoError(t, f.SetCellFormula("Sheet1", "B3", "SUM(1048576:1:2)"))
	assert.Equal(t, ErrMaxRows, f.InsertRows("Sheet1", 1, 1))

	f = NewFile()
	assert.NoError(t, f.SetCellFormula("Sheet1", "B3", "SUM(XFD:A:B)"))
	assert.NoError(t, f.InsertCols("Sheet1", "A", 1))

	f = NewFile()
	assert.NoError(t, f.SetCellFormula("Sheet1", "B3", "SUM(A:B:XFD)"))
	assert.NoError(t, f.InsertCols("Sheet1", "A", 1))

	// Test adjust formula with defined name in formula text
	f = NewFile()
	assert.NoError(t, f.SetDefinedName(&DefinedName{
		Name:     "Amount",
		RefersTo: "Sheet1!$B$2",
	}))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B2", "Amount+B3"))
	assert.NoError(t, f.RemoveRow("Sheet1", 1))
	formula, err := f.GetCellFormula("Sheet1", "B1")
	assert.NoError(t, err)
	assert.Equal(t, "Amount+B2", formula)

	// Test adjust formula with array formula
	f = NewFile()
	formulaType, reference := STCellFormulaTypeArray, "A3:A3"
	assert.NoError(t, f.SetCellFormula("Sheet1", "A3", "A1:A2", FormulaOpts{Ref: &reference, Type: &formulaType}))
	assert.NoError(t, f.InsertRows("Sheet1", 1, 1))
	formula, err = f.GetCellFormula("Sheet1", "A4")
	assert.NoError(t, err)
	assert.Equal(t, "A2:A3", formula)

	// Test adjust formula on duplicate row with array formula
	f = NewFile()
	formulaType, reference = STCellFormulaTypeArray, "A3"
	assert.NoError(t, f.SetCellFormula("Sheet1", "A3", "A1:A2", FormulaOpts{Ref: &reference, Type: &formulaType}))
	assert.NoError(t, f.InsertRows("Sheet1", 1, 1))
	formula, err = f.GetCellFormula("Sheet1", "A4")
	assert.NoError(t, err)
	assert.Equal(t, "A2:A3", formula)

	// Test adjust formula on duplicate row with relative and absolute cell references
	f = NewFile()
	assert.NoError(t, f.SetCellFormula("Sheet1", "B10", "A$10+$A11&\" \""))
	assert.NoError(t, f.DuplicateRowTo("Sheet1", 10, 2))
	formula, err = f.GetCellFormula("Sheet1", "B2")
	assert.NoError(t, err)
	assert.Equal(t, "A$2+$A3&\" \"", formula)

	t.Run("for_cells_affected_directly", func(t *testing.T) {
		// Test insert row in middle of range with relative and absolute cell references
		f := NewFile()
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "$A1+A$2"))
		assert.NoError(t, f.InsertRows("Sheet1", 2, 1))
		formula, err := f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "$A1+A$3", formula)
		assert.NoError(t, f.RemoveRow("Sheet1", 2))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "$A1+A$2", formula)

		// Test insert column in middle of range
		f = NewFile()
		assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "B1+C1"))
		assert.NoError(t, f.InsertCols("Sheet1", "C", 1))
		formula, err = f.GetCellFormula("Sheet1", "A1")
		assert.NoError(t, err)
		assert.Equal(t, "B1+D1", formula)
		assert.NoError(t, f.RemoveCol("Sheet1", "C"))
		formula, err = f.GetCellFormula("Sheet1", "A1")
		assert.NoError(t, err)
		assert.Equal(t, "B1+C1", formula)

		// Test insert row and column in a rectangular range
		f = NewFile()
		assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "D4+D5+E4+E5"))
		assert.NoError(t, f.InsertCols("Sheet1", "E", 1))
		assert.NoError(t, f.InsertRows("Sheet1", 5, 1))
		formula, err = f.GetCellFormula("Sheet1", "A1")
		assert.NoError(t, err)
		assert.Equal(t, "D4+D6+F4+F6", formula)

		// Test insert row in middle of range
		f = NewFile()
		formulaType, reference := STCellFormulaTypeArray, "B1:B1"
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "A1:A2", FormulaOpts{Ref: &reference, Type: &formulaType}))
		assert.NoError(t, f.InsertRows("Sheet1", 2, 1))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "A1:A3", formula)
		assert.NoError(t, f.RemoveRow("Sheet1", 2))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "A1:A2", formula)

		// Test insert column in middle of range
		f = NewFile()
		formulaType, reference = STCellFormulaTypeArray, "A1:A1"
		assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "B1:C1", FormulaOpts{Ref: &reference, Type: &formulaType}))
		assert.NoError(t, f.InsertCols("Sheet1", "C", 1))
		formula, err = f.GetCellFormula("Sheet1", "A1")
		assert.NoError(t, err)
		assert.Equal(t, "B1:D1", formula)
		assert.NoError(t, f.RemoveCol("Sheet1", "C"))
		formula, err = f.GetCellFormula("Sheet1", "A1")
		assert.NoError(t, err)
		assert.Equal(t, "B1:C1", formula)

		// Test insert row and column in a rectangular range
		f = NewFile()
		formulaType, reference = STCellFormulaTypeArray, "A1:A1"
		assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "D4:E5", FormulaOpts{Ref: &reference, Type: &formulaType}))
		assert.NoError(t, f.InsertCols("Sheet1", "E", 1))
		assert.NoError(t, f.InsertRows("Sheet1", 5, 1))
		formula, err = f.GetCellFormula("Sheet1", "A1")
		assert.NoError(t, err)
		assert.Equal(t, "D4:F6", formula)
	})
	t.Run("for_cells_affected_indirectly", func(t *testing.T) {
		f := NewFile()
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "A3+A4"))
		assert.NoError(t, f.InsertRows("Sheet1", 2, 1))
		formula, err := f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "A4+A5", formula)
		assert.NoError(t, f.RemoveRow("Sheet1", 2))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "A3+A4", formula)

		f = NewFile()
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "D3+D4"))
		assert.NoError(t, f.InsertCols("Sheet1", "C", 1))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "E3+E4", formula)
		assert.NoError(t, f.RemoveCol("Sheet1", "C"))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "D3+D4", formula)
	})
	t.Run("for_entire_cols_rows_reference", func(t *testing.T) {
		f := NewFile()
		// Test adjust formula on insert row in the middle of the range
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "SUM(A2:A3:A4,,Table1[])"))
		assert.NoError(t, f.InsertRows("Sheet1", 3, 1))
		formula, err := f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "SUM(A2:A4:A5,,Table1[])", formula)

		// Test adjust formula on insert at the top of the range
		assert.NoError(t, f.InsertRows("Sheet1", 2, 1))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "SUM(A3:A5:A6,,Table1[])", formula)

		f = NewFile()
		// Test adjust formula on insert row in the middle of the range
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "SUM('Sheet 1'!A2,A3)"))
		assert.NoError(t, f.InsertRows("Sheet1", 3, 1))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "SUM('Sheet 1'!A2,A4)", formula)

		// Test adjust formula on insert row at the top of the range
		assert.NoError(t, f.InsertRows("Sheet1", 2, 1))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "SUM('Sheet 1'!A2,A5)", formula)

		f = NewFile()
		// Test adjust formula on insert col in the middle of the range
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "SUM(C3:D3)"))
		assert.NoError(t, f.InsertCols("Sheet1", "D", 1))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "SUM(C3:E3)", formula)

		// Test adjust formula on insert at the top of the range
		assert.NoError(t, f.InsertCols("Sheet1", "C", 1))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "SUM(D3:F3)", formula)

		f = NewFile()
		// Test adjust formula on insert column in the middle of the range
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "SUM(C3,D3)"))
		assert.NoError(t, f.InsertCols("Sheet1", "D", 1))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "SUM(C3,E3)", formula)

		// Test adjust formula on insert column at the top of the range
		assert.NoError(t, f.InsertCols("Sheet1", "C", 1))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "SUM(D3,F3)", formula)

		f = NewFile()
		// Test adjust formula on insert row in the middle of the range (range of whole row)
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "SUM(2:3)"))
		assert.NoError(t, f.InsertRows("Sheet1", 3, 1))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "SUM(2:4)", formula)

		// Test adjust formula on insert row at the top of the range (range of whole row)
		assert.NoError(t, f.InsertRows("Sheet1", 2, 1))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "SUM(3:5)", formula)

		f = NewFile()
		// Test adjust formula on insert row in the middle of the range (range of whole column)
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "SUM(C:D)"))
		assert.NoError(t, f.InsertCols("Sheet1", "D", 1))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "SUM(C:E)", formula)

		// Test adjust formula on insert row at the top of the range (range of whole column)
		assert.NoError(t, f.InsertCols("Sheet1", "C", 1))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "SUM(D:F)", formula)
	})
	t.Run("for_all_worksheet_cells_with_rows_insert", func(t *testing.T) {
		f := NewFile()
		_, err := f.NewSheet("Sheet2")
		assert.NoError(t, err)
		// Tests formulas referencing Sheet2 should update but those referencing the original sheet should not
		tbl := [][]string{
			{"B1", "Sheet2!A1+Sheet2!A2", "Sheet2!A1+Sheet2!A3", "Sheet2!A2+Sheet2!A4"},
			{"C1", "A1+A2", "A1+A2", "A1+A2"},
			{"D1", "Sheet2!B1:B2", "Sheet2!B1:B3", "Sheet2!B2:B4"},
			{"E1", "B1:B2", "B1:B2", "B1:B2"},
			{"F1", "SUM(Sheet2!C1:C2)", "SUM(Sheet2!C1:C3)", "SUM(Sheet2!C2:C4)"},
			{"G1", "SUM(C1:C2)", "SUM(C1:C2)", "SUM(C1:C2)"},
			{"H1", "SUM(Sheet2!D1,Sheet2!D2)", "SUM(Sheet2!D1,Sheet2!D3)", "SUM(Sheet2!D2,Sheet2!D4)"},
			{"I1", "SUM(D1,D2)", "SUM(D1,D2)", "SUM(D1,D2)"},
		}
		for _, preset := range tbl {
			assert.NoError(t, f.SetCellFormula("Sheet1", preset[0], preset[1]))
		}
		// Test adjust formula on insert row in the middle of the range
		assert.NoError(t, f.InsertRows("Sheet2", 2, 1))
		for _, preset := range tbl {
			formula, err := f.GetCellFormula("Sheet1", preset[0])
			assert.NoError(t, err)
			assert.Equal(t, preset[2], formula)
		}

		// Test adjust formula on insert row in the top of the range
		assert.NoError(t, f.InsertRows("Sheet2", 1, 1))
		for _, preset := range tbl {
			formula, err := f.GetCellFormula("Sheet1", preset[0])
			assert.NoError(t, err)
			assert.Equal(t, preset[3], formula)
		}
	})
	t.Run("for_all_worksheet_cells_with_cols_insert", func(t *testing.T) {
		f := NewFile()
		_, err := f.NewSheet("Sheet2")
		assert.NoError(t, err)
		tbl := [][]string{
			{"A1", "Sheet2!A1+Sheet2!B1", "Sheet2!A1+Sheet2!C1", "Sheet2!B1+Sheet2!D1"},
			{"A2", "A1+B1", "A1+B1", "A1+B1"},
			{"A3", "Sheet2!A2:B2", "Sheet2!A2:C2", "Sheet2!B2:D2"},
			{"A4", "A2:B2", "A2:B2", "A2:B2"},
			{"A5", "SUM(Sheet2!A3:B3)", "SUM(Sheet2!A3:C3)", "SUM(Sheet2!B3:D3)"},
			{"A6", "SUM(A3:B3)", "SUM(A3:B3)", "SUM(A3:B3)"},
			{"A7", "SUM(Sheet2!A4,Sheet2!B4)", "SUM(Sheet2!A4,Sheet2!C4)", "SUM(Sheet2!B4,Sheet2!D4)"},
			{"A8", "SUM(A4,B4)", "SUM(A4,B4)", "SUM(A4,B4)"},
		}
		for _, preset := range tbl {
			assert.NoError(t, f.SetCellFormula("Sheet1", preset[0], preset[1]))
		}
		// Test adjust formula on insert column in the middle of the range
		assert.NoError(t, f.InsertCols("Sheet2", "B", 1))
		for _, preset := range tbl {
			formula, err := f.GetCellFormula("Sheet1", preset[0])
			assert.NoError(t, err)
			assert.Equal(t, preset[2], formula)
		}
		// Test adjust formula on insert column in the top of the range
		assert.NoError(t, f.InsertCols("Sheet2", "A", 1))
		for _, preset := range tbl {
			formula, err := f.GetCellFormula("Sheet1", preset[0])
			assert.NoError(t, err)
			assert.Equal(t, preset[3], formula)
		}
	})
	t.Run("for_cross_sheet_ref_with_rows_insert)", func(t *testing.T) {
		f := NewFile()
		_, err := f.NewSheet("Sheet2")
		assert.NoError(t, err)
		_, err = f.NewSheet("Sheet3")
		assert.NoError(t, err)
		// Tests formulas referencing Sheet2 should update but those referencing
		// the original sheet or Sheet 3 should not update
		tbl := [][]string{
			{"B1", "Sheet2!A1+Sheet2!A2+Sheet1!A3+Sheet1!A4", "Sheet2!A1+Sheet2!A3+Sheet1!A3+Sheet1!A4", "Sheet2!A2+Sheet2!A4+Sheet1!A3+Sheet1!A4"},
			{"C1", "Sheet2!B1+Sheet2!B2+B3+B4", "Sheet2!B1+Sheet2!B3+B3+B4", "Sheet2!B2+Sheet2!B4+B3+B4"},
			{"D1", "Sheet2!C1+Sheet2!C2+Sheet3!A3+Sheet3!A4", "Sheet2!C1+Sheet2!C3+Sheet3!A3+Sheet3!A4", "Sheet2!C2+Sheet2!C4+Sheet3!A3+Sheet3!A4"},
			{"E1", "SUM(Sheet2!D1:D2,Sheet1!A3:A4)", "SUM(Sheet2!D1:D3,Sheet1!A3:A4)", "SUM(Sheet2!D2:D4,Sheet1!A3:A4)"},
			{"F1", "SUM(Sheet2!E1:E2,A3:A4)", "SUM(Sheet2!E1:E3,A3:A4)", "SUM(Sheet2!E2:E4,A3:A4)"},
			{"G1", "SUM(Sheet2!F1:F2,Sheet3!A3:A4)", "SUM(Sheet2!F1:F3,Sheet3!A3:A4)", "SUM(Sheet2!F2:F4,Sheet3!A3:A4)"},
		}
		for _, preset := range tbl {
			assert.NoError(t, f.SetCellFormula("Sheet1", preset[0], preset[1]))
		}
		// Test adjust formula on insert row in the middle of the range
		assert.NoError(t, f.InsertRows("Sheet2", 2, 1))
		for _, preset := range tbl {
			formula, err := f.GetCellFormula("Sheet1", preset[0])
			assert.NoError(t, err)
			assert.Equal(t, preset[2], formula)
		}
		// Test adjust formula on insert row in the top of the range
		assert.NoError(t, f.InsertRows("Sheet2", 1, 1))
		for _, preset := range tbl {
			formula, err := f.GetCellFormula("Sheet1", preset[0])
			assert.NoError(t, err)
			assert.Equal(t, preset[3], formula)
		}
	})
	t.Run("for_cross_sheet_ref_with_cols_insert)", func(t *testing.T) {
		f := NewFile()
		_, err := f.NewSheet("Sheet2")
		assert.NoError(t, err)
		_, err = f.NewSheet("Sheet3")
		assert.NoError(t, err)
		// Tests formulas referencing Sheet2 should update but those referencing
		// the original sheet or Sheet 3 should not update
		tbl := [][]string{
			{"A1", "Sheet2!A1+Sheet2!B1+Sheet1!C1+Sheet1!D1", "Sheet2!A1+Sheet2!C1+Sheet1!C1+Sheet1!D1", "Sheet2!B1+Sheet2!D1+Sheet1!C1+Sheet1!D1"},
			{"A2", "Sheet2!A2+Sheet2!B2+C2+D2", "Sheet2!A2+Sheet2!C2+C2+D2", "Sheet2!B2+Sheet2!D2+C2+D2"},
			{"A3", "Sheet2!A3+Sheet2!B3+Sheet3!C3+Sheet3!D3", "Sheet2!A3+Sheet2!C3+Sheet3!C3+Sheet3!D3", "Sheet2!B3+Sheet2!D3+Sheet3!C3+Sheet3!D3"},
			{"A4", "SUM(Sheet2!A4:B4,Sheet1!C4:D4)", "SUM(Sheet2!A4:C4,Sheet1!C4:D4)", "SUM(Sheet2!B4:D4,Sheet1!C4:D4)"},
			{"A5", "SUM(Sheet2!A5:B5,C5:D5)", "SUM(Sheet2!A5:C5,C5:D5)", "SUM(Sheet2!B5:D5,C5:D5)"},
			{"A6", "SUM(Sheet2!A6:B6,Sheet3!C6:D6)", "SUM(Sheet2!A6:C6,Sheet3!C6:D6)", "SUM(Sheet2!B6:D6,Sheet3!C6:D6)"},
		}
		for _, preset := range tbl {
			assert.NoError(t, f.SetCellFormula("Sheet1", preset[0], preset[1]))
		}
		// Test adjust formula on insert row in the middle of the range
		assert.NoError(t, f.InsertCols("Sheet2", "B", 1))
		for _, preset := range tbl {
			formula, err := f.GetCellFormula("Sheet1", preset[0])
			assert.NoError(t, err)
			assert.Equal(t, preset[2], formula)
		}
		// Test adjust formula on insert row in the top of the range
		assert.NoError(t, f.InsertCols("Sheet2", "A", 1))
		for _, preset := range tbl {
			formula, err := f.GetCellFormula("Sheet1", preset[0])
			assert.NoError(t, err)
			assert.Equal(t, preset[3], formula)
		}
	})
	t.Run("for_cross_sheet_ref_with_chart_sheet)", func(t *testing.T) {
		assert.NoError(t, f.AddChartSheet("Chart1", &Chart{Type: Line}))
		assert.NoError(t, f.InsertRows("Sheet1", 2, 1))
		assert.NoError(t, f.InsertCols("Sheet1", "A", 1))
	})
	t.Run("for_array_formula_cell", func(t *testing.T) {
		f := NewFile()
		assert.NoError(t, f.SetSheetRow("Sheet1", "A1", &[]int{1, 2}))
		assert.NoError(t, f.SetSheetRow("Sheet1", "A2", &[]int{3, 4}))
		formulaType, ref := STCellFormulaTypeArray, "C1:C2"
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", "A1:A2*B1:B2", FormulaOpts{Ref: &ref, Type: &formulaType}))
		assert.NoError(t, f.InsertRows("Sheet1", 1, 1))
		assert.NoError(t, f.InsertCols("Sheet1", "A", 1))
		result, err := f.CalcCellValue("Sheet1", "D2")
		assert.NoError(t, err)
		assert.Equal(t, "2", result)
		result, err = f.CalcCellValue("Sheet1", "D3")
		assert.NoError(t, err)
		assert.Equal(t, "12", result)

		// Test adjust array formula with column overflow produces #REF!
		formulaType, ref = STCellFormulaTypeArray, "E1:E2"
		assert.NoError(t, f.SetCellFormula("Sheet1", "E1", "XFD1:XFD1", FormulaOpts{Ref: &ref, Type: &formulaType}))
		assert.NoError(t, f.InsertCols("Sheet1", "A", 1))
	})
}

func TestAdjustVolatileDeps(t *testing.T) {
	f := NewFile()
	f.Pkg.Store(defaultXMLPathVolatileDeps, []byte(fmt.Sprintf(`<volTypes xmlns="%s"><volType><main><tp><tr r="C2" s="2"/><tr r="C2" s="1"/><tr r="D3" s="1"/></tp></main></volType></volTypes>`, NameSpaceSpreadSheet.Value)))
	assert.NoError(t, f.InsertCols("Sheet1", "A", 1))
	assert.NoError(t, f.InsertRows("Sheet1", 2, 1))
	assert.Equal(t, "D3", f.VolatileDeps.VolType[0].Main[0].Tp[0].Tr[1].R)
	assert.NoError(t, f.RemoveCol("Sheet1", "D"))
	assert.NoError(t, f.RemoveRow("Sheet1", 4))
	assert.Len(t, f.VolatileDeps.VolType[0].Main[0].Tp[0].Tr, 1)

	f = NewFile()
	f.Pkg.Store(defaultXMLPathVolatileDeps, MacintoshCyrillicCharset)
	assert.EqualError(t, f.InsertRows("Sheet1", 2, 1), "XML syntax error on line 1: invalid UTF-8")

	f = NewFile()
	f.Pkg.Store(defaultXMLPathVolatileDeps, []byte(fmt.Sprintf(`<volTypes xmlns="%s"><volType><main><tp><tr r="A" s="1"/></tp></main></volType></volTypes>`, NameSpaceSpreadSheet.Value)))
	assert.Equal(t, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")), f.InsertCols("Sheet1", "A", 1))
	f.volatileDepsWriter()
}

func TestAdjustConditionalFormats(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetSheetRow("Sheet1", "B1", &[]interface{}{1, nil, 1, 1}))
	formatID, err := f.NewConditionalStyle(&Style{Font: &Font{Color: "09600B"}, Fill: Fill{Type: "pattern", Color: []string{"C7EECF"}, Pattern: 1}})
	assert.NoError(t, err)
	format := []ConditionalFormatOptions{
		{
			Type:     "cell",
			Criteria: "greater than",
			Format:   &formatID,
			Value:    "0",
		},
	}
	for _, ref := range []string{"B1", "D1:E1"} {
		assert.NoError(t, f.SetConditionalFormat("Sheet1", ref, format))
	}
	assert.NoError(t, f.RemoveCol("Sheet1", "B"))
	opts, err := f.GetConditionalFormats("Sheet1")
	assert.NoError(t, err)
	assert.Len(t, format, 1)
	assert.Equal(t, format, opts["C1:D1"])

	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).ConditionalFormatting[0].SQRef = "-"
	assert.Equal(t, newCellNameToCoordinatesError("-", newInvalidCellNameError("-")), f.RemoveCol("Sheet1", "B"))

	ws.(*xlsxWorksheet).ConditionalFormatting[0] = nil
	assert.NoError(t, f.RemoveCol("Sheet1", "B"))

	t.Run("for_remove_conditional_formats_column", func(t *testing.T) {
		f := NewFile()
		format := []ConditionalFormatOptions{{
			Type:     "data_bar",
			Criteria: "=",
			MinType:  "min",
			MaxType:  "max",
			BarColor: "#638EC6",
		}}
		assert.NoError(t, f.SetConditionalFormat("Sheet1", "D2:D3", format))
		assert.NoError(t, f.SetConditionalFormat("Sheet1", "D5", format))
		assert.NoError(t, f.RemoveCol("Sheet1", "D"))
		opts, err := f.GetConditionalFormats("Sheet1")
		assert.NoError(t, err)
		assert.Len(t, opts, 0)
	})
	t.Run("for_remove_conditional_formats_row", func(t *testing.T) {
		f := NewFile()
		format := []ConditionalFormatOptions{{
			Type:     "data_bar",
			Criteria: "=",
			MinType:  "min",
			MaxType:  "max",
			BarColor: "#638EC6",
		}}
		assert.NoError(t, f.SetConditionalFormat("Sheet1", "D2:E2", format))
		assert.NoError(t, f.SetConditionalFormat("Sheet1", "F2", format))
		assert.NoError(t, f.RemoveRow("Sheet1", 2))
		opts, err := f.GetConditionalFormats("Sheet1")
		assert.NoError(t, err)
		assert.Len(t, opts, 0)
	})
	t.Run("for_adjust_conditional_formats_row", func(t *testing.T) {
		f := NewFile()
		format := []ConditionalFormatOptions{{
			Type:     "data_bar",
			Criteria: "=",
			MinType:  "min",
			MaxType:  "max",
			BarColor: "#638EC6",
		}}
		assert.NoError(t, f.SetConditionalFormat("Sheet1", "D2:D3", format))
		assert.NoError(t, f.SetConditionalFormat("Sheet1", "D5", format))
		assert.NoError(t, f.RemoveRow("Sheet1", 1))
		opts, err := f.GetConditionalFormats("Sheet1")
		assert.NoError(t, err)
		assert.Len(t, opts, 2)
		assert.Equal(t, format, opts["D1:D2"])
		assert.Equal(t, format, opts["D4:D4"])
	})
}

func TestAdjustDataValidations(t *testing.T) {
	f := NewFile()
	dv := NewDataValidation(true)
	dv.Sqref = "B1"
	assert.NoError(t, dv.SetDropList([]string{"1", "2", "3"}))
	assert.NoError(t, f.AddDataValidation("Sheet1", dv))
	assert.NoError(t, f.RemoveCol("Sheet1", "B"))
	dvs, err := f.GetDataValidations("Sheet1")
	assert.NoError(t, err)
	assert.Len(t, dvs, 0)

	assert.NoError(t, f.SetCellValue("Sheet1", "F2", 1))
	assert.NoError(t, f.SetCellValue("Sheet1", "F3", 2))
	dv = NewDataValidation(true)
	dv.Sqref = "C2:D3"
	dv.SetSqrefDropList("$F$2:$F$3")
	assert.NoError(t, f.AddDataValidation("Sheet1", dv))

	assert.NoError(t, f.AddChartSheet("Chart1", &Chart{Type: Line}))
	_, err = f.NewSheet("Sheet2")
	assert.NoError(t, err)
	assert.NoError(t, f.SetSheetRow("Sheet2", "C1", &[]interface{}{1, 10}))
	dv = NewDataValidation(true)
	dv.Sqref = "C5:D6"
	assert.NoError(t, dv.SetRange("Sheet2!C1", "Sheet2!D1", DataValidationTypeWhole, DataValidationOperatorBetween))
	dv.SetError(DataValidationErrorStyleStop, "error title", "error body")
	assert.NoError(t, f.AddDataValidation("Sheet1", dv))
	assert.NoError(t, f.RemoveCol("Sheet1", "B"))
	assert.NoError(t, f.RemoveCol("Sheet2", "B"))
	dvs, err = f.GetDataValidations("Sheet1")
	assert.NoError(t, err)
	assert.Equal(t, "B2:C3", dvs[0].Sqref)
	assert.Equal(t, "$E$2:$E$3", dvs[0].Formula1)
	assert.Equal(t, "B5:C6", dvs[1].Sqref)
	assert.Equal(t, "Sheet2!B1", dvs[1].Formula1)
	assert.Equal(t, "Sheet2!C1", dvs[1].Formula2)

	dv = NewDataValidation(true)
	dv.Sqref = "C8:D10"
	assert.NoError(t, dv.SetDropList([]string{`A<`, `B>`, `C"`, "D\t", `E'`, `F`}))
	assert.NoError(t, f.AddDataValidation("Sheet1", dv))
	assert.NoError(t, f.RemoveCol("Sheet1", "B"))
	dvs, err = f.GetDataValidations("Sheet1")
	assert.NoError(t, err)
	assert.Equal(t, "\"A<,B>,C\",D\t,E',F\"", dvs[2].Formula1)

	// Test adjust data validation with multiple cell range
	dv = NewDataValidation(true)
	dv.Sqref = "G1:G3 H1:H3 A3:A1048576"
	assert.NoError(t, dv.SetDropList([]string{"1", "2", "3"}))
	assert.NoError(t, f.AddDataValidation("Sheet1", dv))
	assert.NoError(t, f.InsertRows("Sheet1", 2, 1))
	dvs, err = f.GetDataValidations("Sheet1")
	assert.NoError(t, err)
	assert.Equal(t, "G1:G4 H1:H4 A4:A1048576", dvs[3].Sqref)

	dv = NewDataValidation(true)
	dv.Sqref = "C5:D6"
	assert.NoError(t, dv.SetRange("Sheet1!A1048576", "Sheet1!XFD1", DataValidationTypeWhole, DataValidationOperatorBetween))
	dv.SetError(DataValidationErrorStyleStop, "error title", "error body")
	assert.NoError(t, f.AddDataValidation("Sheet1", dv))
	assert.NoError(t, f.InsertCols("Sheet1", "A", 1))
	assert.Equal(t, ErrMaxRows, f.InsertRows("Sheet1", 1, 1))

	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).DataValidations.DataValidation[0].Sqref = "-"
	assert.Equal(t, newCellNameToCoordinatesError("-", newInvalidCellNameError("-")), f.RemoveCol("Sheet1", "B"))

	ws.(*xlsxWorksheet).DataValidations.DataValidation[0] = nil
	assert.NoError(t, f.RemoveCol("Sheet1", "B"))

	ws.(*xlsxWorksheet).DataValidations = nil
	assert.NoError(t, f.RemoveCol("Sheet1", "B"))

	f.Sheet.Delete("xl/worksheets/sheet1.xml")
	f.Pkg.Store("xl/worksheets/sheet1.xml", MacintoshCyrillicCharset)
	assert.EqualError(t, f.adjustDataValidations(nil, "Sheet1", columns, 0, 0, 1), "XML syntax error on line 1: invalid UTF-8")

	t.Run("for_escaped_data_validation_rules_formula", func(t *testing.T) {
		f := NewFile()
		_, err := f.NewSheet("Sheet2")
		assert.NoError(t, err)
		dv := NewDataValidation(true)
		dv.Sqref = "A1"
		assert.NoError(t, dv.SetDropList([]string{"option1", strings.Repeat("\"", 4)}))
		ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
		assert.True(t, ok)
		assert.NoError(t, f.AddDataValidation("Sheet1", dv))
		// The double quote symbol in none formula data validation rules will be escaped in the Kingsoft WPS Office
		formula := strings.ReplaceAll(fmt.Sprintf("\"option1, %s", strings.Repeat("\"", 9)), "\"", "&quot;")
		ws.(*xlsxWorksheet).DataValidations.DataValidation[0].Formula1.Content = formula
		assert.NoError(t, f.RemoveCol("Sheet2", "A"))
		dvs, err := f.GetDataValidations("Sheet1")
		assert.NoError(t, err)
		assert.Equal(t, formula, dvs[0].Formula1)
	})

	t.Run("no_data_validations_on_first_sheet", func(t *testing.T) {
		f := NewFile()

		// Add Sheet2 and set a data validation
		_, err = f.NewSheet("Sheet2")
		assert.NoError(t, err)
		dv := NewDataValidation(true)
		dv.Sqref = "C5:D6"
		assert.NoError(t, f.AddDataValidation("Sheet2", dv))

		// Adjust Sheet2 by removing a column
		assert.NoError(t, f.RemoveCol("Sheet2", "A"))

		// Verify that data validations on Sheet2 are adjusted correctly
		dvs, err = f.GetDataValidations("Sheet2")
		assert.NoError(t, err)
		assert.Equal(t, "B5:C6", dvs[0].Sqref) // Adjusted range
	})
}

func TestAdjustDrawings(t *testing.T) {
	f := NewFile()
	// Test add pictures to sheet with positioning
	assert.NoError(t, f.AddPicture("Sheet1", "B2", filepath.Join("test", "images", "excel.jpg"), nil))
	assert.NoError(t, f.AddPicture("Sheet1", "B11", filepath.Join("test", "images", "excel.jpg"), &GraphicOptions{Positioning: "oneCell"}))
	assert.NoError(t, f.AddPicture("Sheet1", "B21", filepath.Join("test", "images", "excel.jpg"), &GraphicOptions{Positioning: "absolute"}))

	// Test adjust pictures on inserting columns and rows
	assert.NoError(t, f.InsertCols("Sheet1", "A", 1))
	assert.NoError(t, f.InsertRows("Sheet1", 1, 1))
	assert.NoError(t, f.InsertCols("Sheet1", "C", 1))
	assert.NoError(t, f.InsertRows("Sheet1", 5, 1))
	assert.NoError(t, f.InsertRows("Sheet1", 15, 1))
	cells, err := f.GetPictureCells("Sheet1")
	assert.NoError(t, err)
	assert.Equal(t, []string{"D3", "B21", "D13"}, cells)
	wb := filepath.Join("test", "TestAdjustDrawings.xlsx")
	assert.NoError(t, f.SaveAs(wb))

	// Test adjust pictures on deleting columns and rows
	assert.NoError(t, f.RemoveCol("Sheet1", "A"))
	assert.NoError(t, f.RemoveRow("Sheet1", 1))
	cells, err = f.GetPictureCells("Sheet1")
	assert.NoError(t, err)
	assert.Equal(t, []string{"C2", "B21", "C12"}, cells)

	// Test adjust existing pictures on inserting columns and rows
	f, err = OpenFile(wb)
	assert.NoError(t, err)
	assert.NoError(t, f.InsertCols("Sheet1", "A", 1))
	assert.NoError(t, f.InsertRows("Sheet1", 1, 1))
	assert.NoError(t, f.InsertCols("Sheet1", "D", 1))
	assert.NoError(t, f.InsertRows("Sheet1", 5, 1))
	assert.NoError(t, f.InsertRows("Sheet1", 16, 1))
	cells, err = f.GetPictureCells("Sheet1")
	assert.NoError(t, err)
	assert.Equal(t, []string{"F4", "B21", "F15"}, cells)

	// Test adjust drawings with unsupported charset
	f, err = OpenFile(wb)
	assert.NoError(t, err)
	f.Pkg.Store("xl/drawings/drawing1.xml", MacintoshCyrillicCharset)
	assert.EqualError(t, f.InsertCols("Sheet1", "A", 1), "XML syntax error on line 1: invalid UTF-8")

	// Drawings at boundary columns now cap to MaxColumns-1 instead of erroring
	cells = []string{"XFD1", "XFB1"}
	for _, cell := range cells {
		f = NewFile()
		assert.NoError(t, f.AddPicture("Sheet1", cell, filepath.Join("test", "images", "excel.jpg"), nil))
		assert.NoError(t, f.InsertCols("Sheet1", "A", 1))
		assert.NoError(t, f.SaveAs(wb))
		f, err = OpenFile(wb)
		assert.NoError(t, err)
		assert.NoError(t, f.InsertCols("Sheet1", "A", 1))
	}
	rowErrors := []error{ErrMaxRows, ErrMaxRows}
	cells = []string{"A1048576", "A1048570"}
	for i, cell := range cells {
		f = NewFile()
		assert.NoError(t, f.AddPicture("Sheet1", cell, filepath.Join("test", "images", "excel.jpg"), nil))
		assert.Equal(t, rowErrors[i], f.InsertRows("Sheet1", 1, 1))
		assert.NoError(t, f.SaveAs(wb))
		f, err = OpenFile(wb)
		assert.NoError(t, err)
		assert.Equal(t, rowErrors[i], f.InsertRows("Sheet1", 1, 1))
	}

	a := xdrCellAnchor{}
	assert.NoError(t, a.adjustDrawings(columns, 0, 0))
	p := xlsxCellAnchorPos{}
	assert.NoError(t, p.adjustDrawings(columns, 0, 0, ""))

	f, err = OpenFile(wb)
	assert.NoError(t, err)
	f.Pkg.Store("xl/drawings/drawing1.xml", []byte(xml.Header+`<wsDr xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><twoCellAnchor><from><col>0</col><colOff>0</colOff><row>0</row><rowOff>0</rowOff></from><to><col>1</col><colOff>0</colOff><row>1</row><rowOff>0</rowOff></to><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"></mc:AlternateContent><clientData/></twoCellAnchor></wsDr>`))
	assert.NoError(t, f.InsertCols("Sheet1", "A", 1))

	f, err = OpenFile(wb)
	assert.NoError(t, err)
	f.Pkg.Store("xl/drawings/drawing1.xml", []byte(xml.Header+fmt.Sprintf(`<wsDr xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><oneCellAnchor><from><col>%d</col><row>0</row></from><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"></mc:AlternateContent><clientData/></oneCellAnchor></wsDr>`, MaxColumns)))
	// Drawing column at MaxColumns now caps to MaxColumns-1 instead of erroring
	assert.NoError(t, f.InsertCols("Sheet1", "A", 1))
}

func TestAdjustDefinedNames(t *testing.T) {
	f := NewFile()
	_, err := f.NewSheet("Sheet2")
	assert.NoError(t, err)
	for _, dn := range []*DefinedName{
		{Name: "Name1", RefersTo: "Sheet1!$XFD$1"},
		{Name: "Name2", RefersTo: "Sheet2!$C$1", Scope: "Sheet1"},
		{Name: "Name3", RefersTo: "Sheet2!$C$1:$D$2", Scope: "Sheet1"},
		{Name: "Name4", RefersTo: "Sheet2!$C1:D$2"},
		{Name: "Name5", RefersTo: "Sheet2!C$1:$D2"},
		{Name: "Name6", RefersTo: "Sheet2!C:$D"},
		{Name: "Name7", RefersTo: "Sheet2!$C:D"},
		{Name: "Name8", RefersTo: "Sheet2!C:D"},
		{Name: "Name9", RefersTo: "Sheet2!$C:$D"},
		{Name: "Name10", RefersTo: "Sheet2!1:2"},
	} {
		assert.NoError(t, f.SetDefinedName(dn))
	}
	assert.NoError(t, f.InsertCols("Sheet1", "A", 1))
	assert.NoError(t, f.InsertRows("Sheet1", 1, 1))
	assert.NoError(t, f.InsertCols("Sheet2", "A", 1))
	assert.NoError(t, f.InsertRows("Sheet2", 1, 1))
	definedNames := f.GetDefinedName()
	for i, expected := range []string{
		"Sheet1!$XFD$2",
		"Sheet2!$D$2",
		"Sheet2!$D$2:$E$3",
		"Sheet2!$D1:D$3",
		"Sheet2!C$2:$E2",
		"Sheet2!C:$E",
		"Sheet2!$D:D",
		"Sheet2!C:D",
		"Sheet2!$D:$E",
		"Sheet2!1:2",
	} {
		assert.Equal(t, expected, definedNames[i].RefersTo)
	}

	f = NewFile()
	assert.NoError(t, f.SetDefinedName(&DefinedName{
		Name:     "Name1",
		RefersTo: "Sheet1!$A$1",
		Scope:    "Sheet1",
	}))
	assert.NoError(t, f.RemoveCol("Sheet1", "A"))
	definedNames = f.GetDefinedName()
	assert.Equal(t, "Sheet1!$A$1", definedNames[0].RefersTo)

	f = NewFile()
	assert.NoError(t, f.SetDefinedName(&DefinedName{
		Name:     "Name1",
		RefersTo: "'1.A & B C'!#REF!",
		Scope:    "Sheet1",
	}))
	assert.NoError(t, f.RemoveCol("Sheet1", "A"))
	definedNames = f.GetDefinedName()
	assert.Equal(t, "'1.A & B C'!#REF!", definedNames[0].RefersTo)

	f = NewFile()
	f.WorkBook = nil
	f.Pkg.Store(defaultXMLPathWorkbook, MacintoshCyrillicCharset)
	assert.EqualError(t, f.adjustDefinedNames("Sheet1", columns, 0, 0), "XML syntax error on line 1: invalid UTF-8")
}

func TestOnFormulaAdjustedCallback(t *testing.T) {
	t.Run("InsertRows", func(t *testing.T) {
		f := NewFile()
		f.SetCellFormula("Sheet1", "B5", "SUM(A1:A4)")

		var changes []struct{ Sheet, Cell, Old, New string }
		f.OnFormulaAdjusted = func(sheet, cell, oldFormula, newFormula string) {
			changes = append(changes, struct{ Sheet, Cell, Old, New string }{sheet, cell, oldFormula, newFormula})
		}

		assert.NoError(t, f.InsertRows("Sheet1", 2, 1))

		assert.True(t, len(changes) > 0, "expected at least one callback")
		found := false
		for _, c := range changes {
			if c.Sheet == "Sheet1" && c.Cell == "B6" && c.Old == "SUM(A1:A4)" && c.New == "SUM(A1:A5)" {
				found = true
			}
		}
		assert.True(t, found, "expected B6 formula adjusted from SUM(A1:A4) to SUM(A1:A5)")
	})

	t.Run("InsertCols", func(t *testing.T) {
		f := NewFile()
		f.SetCellFormula("Sheet1", "D1", "SUM(A1:C1)")

		var changes []struct{ Sheet, Cell, Old, New string }
		f.OnFormulaAdjusted = func(sheet, cell, oldFormula, newFormula string) {
			changes = append(changes, struct{ Sheet, Cell, Old, New string }{sheet, cell, oldFormula, newFormula})
		}

		assert.NoError(t, f.InsertCols("Sheet1", "B", 1))

		found := false
		for _, c := range changes {
			if c.Sheet == "Sheet1" && c.Old == "SUM(A1:C1)" && c.New == "SUM(A1:D1)" {
				found = true
			}
		}
		assert.True(t, found, "expected formula adjusted from SUM(A1:C1) to SUM(A1:D1)")
	})

	t.Run("NilCallback", func(t *testing.T) {
		f := NewFile()
		f.SetCellFormula("Sheet1", "B5", "SUM(A1:A4)")
		f.OnFormulaAdjusted = nil
		assert.NoError(t, f.InsertRows("Sheet1", 2, 1))
	})

	t.Run("NoChangeNoCallback", func(t *testing.T) {
		f := NewFile()
		f.SetCellFormula("Sheet1", "B1", "SUM(C1:D1)")

		callCount := 0
		f.OnFormulaAdjusted = func(sheet, cell, oldFormula, newFormula string) {
			if cell == "B1" {
				callCount++
			}
		}

		// 在第 5 行插入，不影响 B1 的公式 SUM(C1:D1)
		assert.NoError(t, f.InsertRows("Sheet1", 5, 1))
		assert.Equal(t, 0, callCount, "should not trigger callback when formula unchanged")
	})

	t.Run("DeleteRows", func(t *testing.T) {
		f := NewFile()
		f.SetCellFormula("Sheet1", "B5", "SUM(A1:A4)")

		var changes []struct{ Sheet, Cell, Old, New string }
		f.OnFormulaAdjusted = func(sheet, cell, oldFormula, newFormula string) {
			changes = append(changes, struct{ Sheet, Cell, Old, New string }{sheet, cell, oldFormula, newFormula})
		}

		assert.NoError(t, f.RemoveRow("Sheet1", 1))

		assert.True(t, len(changes) > 0, "expected callback on row deletion")
	})

	t.Run("RenameSheet", func(t *testing.T) {
		f := NewFile()
		f.NewSheet("Data")
		f.SetCellFormula("Sheet1", "A1", "Data!A1*2")
		f.SetCellFormula("Sheet1", "A2", "SUM(Data!A1:A10)")

		var changes []struct{ Sheet, Cell, Old, New string }
		f.OnFormulaAdjusted = func(sheet, cell, oldFormula, newFormula string) {
			changes = append(changes, struct{ Sheet, Cell, Old, New string }{sheet, cell, oldFormula, newFormula})
		}

		assert.NoError(t, f.SetSheetName("Data", "数据表"))

		assert.True(t, len(changes) >= 2, "expected at least 2 callbacks for renamed sheet references")
		foundA1 := false
		foundA2 := false
		for _, c := range changes {
			if c.Sheet == "Sheet1" && c.Cell == "A1" && c.Old == "Data!A1*2" && c.New == "数据表!A1*2" {
				foundA1 = true
			}
			if c.Sheet == "Sheet1" && c.Cell == "A2" && c.Old == "SUM(Data!A1:A10)" && c.New == "SUM(数据表!A1:A10)" {
				foundA2 = true
			}
		}
		assert.True(t, foundA1, "expected A1 formula adjusted from Data!A1*2 to 数据表!A1*2")
		assert.True(t, foundA2, "expected A2 formula adjusted from SUM(Data!A1:A10) to SUM(数据表!A1:A10)")
	})

	t.Run("MoveRow", func(t *testing.T) {
		f := NewFile()
		// A1 引用 A3，把第 3 行移到第 5 行后，A1 的公式应该从 A3 变成 A5
		f.SetCellValue("Sheet1", "A3", 100)
		f.SetCellFormula("Sheet1", "A1", "A3*2")

		var changes []struct{ Sheet, Cell, Old, New string }
		f.OnFormulaAdjusted = func(sheet, cell, oldFormula, newFormula string) {
			changes = append(changes, struct{ Sheet, Cell, Old, New string }{sheet, cell, oldFormula, newFormula})
		}

		assert.NoError(t, f.MoveRow("Sheet1", 3, 5))

		found := false
		for _, c := range changes {
			if c.Sheet == "Sheet1" && c.Old != c.New {
				found = true
			}
		}
		assert.True(t, found, "expected callback on MoveRow formula adjustment")
	})

	t.Run("MoveCol", func(t *testing.T) {
		f := NewFile()
		f.SetCellValue("Sheet1", "C1", 100)
		f.SetCellFormula("Sheet1", "A1", "C1*2")

		var changes []struct{ Sheet, Cell, Old, New string }
		f.OnFormulaAdjusted = func(sheet, cell, oldFormula, newFormula string) {
			changes = append(changes, struct{ Sheet, Cell, Old, New string }{sheet, cell, oldFormula, newFormula})
		}

		assert.NoError(t, f.MoveCol("Sheet1", "C", "E"))

		found := false
		for _, c := range changes {
			if c.Sheet == "Sheet1" && c.Old != c.New {
				found = true
			}
		}
		assert.True(t, found, "expected callback on MoveCol formula adjustment")
	})

	t.Run("MoveRows", func(t *testing.T) {
		f := NewFile()
		f.SetCellValue("Sheet1", "A5", 100)
		f.SetCellFormula("Sheet1", "A1", "A5+A6")

		var changes []struct{ Sheet, Cell, Old, New string }
		f.OnFormulaAdjusted = func(sheet, cell, oldFormula, newFormula string) {
			changes = append(changes, struct{ Sheet, Cell, Old, New string }{sheet, cell, oldFormula, newFormula})
		}

		assert.NoError(t, f.MoveRows("Sheet1", 5, 2, 10))

		found := false
		for _, c := range changes {
			if c.Sheet == "Sheet1" && c.Old != c.New {
				found = true
			}
		}
		assert.True(t, found, "expected callback on MoveRows formula adjustment")
	})

	t.Run("MoveCols", func(t *testing.T) {
		f := NewFile()
		f.SetCellValue("Sheet1", "E1", 100)
		f.SetCellFormula("Sheet1", "A1", "E1+F1")

		var changes []struct{ Sheet, Cell, Old, New string }
		f.OnFormulaAdjusted = func(sheet, cell, oldFormula, newFormula string) {
			changes = append(changes, struct{ Sheet, Cell, Old, New string }{sheet, cell, oldFormula, newFormula})
		}

		assert.NoError(t, f.MoveCols("Sheet1", "E", 2, "H"))

		found := false
		for _, c := range changes {
			if c.Sheet == "Sheet1" && c.Old != c.New {
				found = true
			}
		}
		assert.True(t, found, "expected callback on MoveCols formula adjustment")
	})
}

func TestRefErrorCellValueUpdate(t *testing.T) {
	// helper to get a cell pointer for setting V/T directly in tests
	setCellV := func(f *File, sheet, cellName, v, typ string) {
		ws, err := f.workSheetReader(sheet)
		assert.NoError(t, err)
		c, _, _, err := ws.prepareCell(cellName)
		assert.NoError(t, err)
		c.V = v
		c.T = typ
	}

	t.Run("RemoveRow_causes_REF", func(t *testing.T) {
		f := NewFile()
		// A1 references A5, removing row 5 should produce #REF!
		assert.NoError(t, f.SetCellValue("Sheet1", "A5", 100))
		assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "A5*2"))
		// Pre-set a cached value to verify it gets replaced
		setCellV(f, "Sheet1", "A1", "200", "")

		var calcChanges []struct{ Sheet, Cell, Old, New string }
		f.OnCellCalculated = func(sheet, cell, oldValue, newValue string) {
			calcChanges = append(calcChanges, struct{ Sheet, Cell, Old, New string }{sheet, cell, oldValue, newValue})
		}

		assert.NoError(t, f.RemoveRow("Sheet1", 5))

		// cell.V should now be #REF!
		val, err := f.GetCellValue("Sheet1", "A1")
		assert.NoError(t, err)
		assert.Equal(t, "#REF!", val)

		// OnCellCalculated should have been triggered
		found := false
		for _, c := range calcChanges {
			if c.Sheet == "Sheet1" && c.Cell == "A1" && c.Old == "200" && c.New == "#REF!" {
				found = true
			}
		}
		assert.True(t, found, "expected OnCellCalculated for A1 with old=200 new=#REF!")
	})

	t.Run("RemoveCol_causes_REF", func(t *testing.T) {
		f := NewFile()
		// A1 references E1, removing column E should produce #REF!
		assert.NoError(t, f.SetCellValue("Sheet1", "E1", 50))
		assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "E1+10"))
		setCellV(f, "Sheet1", "A1", "60", "")

		var calcChanges []struct{ Sheet, Cell, Old, New string }
		f.OnCellCalculated = func(sheet, cell, oldValue, newValue string) {
			calcChanges = append(calcChanges, struct{ Sheet, Cell, Old, New string }{sheet, cell, oldValue, newValue})
		}

		assert.NoError(t, f.RemoveCol("Sheet1", "E"))

		val, err := f.GetCellValue("Sheet1", "A1")
		assert.NoError(t, err)
		assert.Equal(t, "#REF!", val)

		found := false
		for _, c := range calcChanges {
			if c.Sheet == "Sheet1" && c.Cell == "A1" && c.Old == "60" && c.New == "#REF!" {
				found = true
			}
		}
		assert.True(t, found, "expected OnCellCalculated for A1 with old=60 new=#REF!")
	})

	t.Run("InsertRow_no_REF", func(t *testing.T) {
		f := NewFile()
		assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "SUM(B1:B5)"))

		refTriggered := false
		f.OnCellCalculated = func(sheet, cell, oldValue, newValue string) {
			if newValue == "#REF!" {
				refTriggered = true
			}
		}

		assert.NoError(t, f.InsertRows("Sheet1", 2, 1))
		assert.False(t, refTriggered, "InsertRows should not produce #REF! callback")
	})

	t.Run("AlreadyHasREF_no_duplicate", func(t *testing.T) {
		f := NewFile()
		// Formula already contains #REF! before deletion
		assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "#REF!+A5"))
		setCellV(f, "Sheet1", "A1", "#REF!", "e")

		callCount := 0
		f.OnCellCalculated = func(sheet, cell, oldValue, newValue string) {
			callCount++
		}

		// Remove row 5 — A5 reference becomes #REF!, but formula already had #REF!
		assert.NoError(t, f.RemoveRow("Sheet1", 5))
		assert.Equal(t, 0, callCount, "should not trigger OnCellCalculated when formula already had #REF!")
	})

	t.Run("CrossSheet_REF", func(t *testing.T) {
		f := NewFile()
		_, err := f.NewSheet("Sheet2")
		assert.NoError(t, err)
		assert.NoError(t, f.SetCellValue("Sheet1", "A5", 999))
		assert.NoError(t, f.SetCellFormula("Sheet2", "B1", "Sheet1!A5*3"))
		setCellV(f, "Sheet2", "B1", "2997", "")

		var calcChanges []struct{ Sheet, Cell, Old, New string }
		f.OnCellCalculated = func(sheet, cell, oldValue, newValue string) {
			calcChanges = append(calcChanges, struct{ Sheet, Cell, Old, New string }{sheet, cell, oldValue, newValue})
		}

		assert.NoError(t, f.RemoveRow("Sheet1", 5))

		// Sheet2 B1 should have #REF! and callback should reference Sheet2
		val, err := f.GetCellValue("Sheet2", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "#REF!", val)

		found := false
		for _, c := range calcChanges {
			if c.Sheet == "Sheet2" && c.Cell == "B1" && c.Old == "2997" && c.New == "#REF!" {
				found = true
			}
		}
		assert.True(t, found, "expected OnCellCalculated for Sheet2!B1")
	})

	t.Run("NilCallback_still_updates_value", func(t *testing.T) {
		f := NewFile()
		assert.NoError(t, f.SetCellValue("Sheet1", "A5", 100))
		assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "A5*2"))
		setCellV(f, "Sheet1", "A1", "200", "")
		f.OnCellCalculated = nil

		assert.NoError(t, f.RemoveRow("Sheet1", 5))

		val, err := f.GetCellValue("Sheet1", "A1")
		assert.NoError(t, err)
		assert.Equal(t, "#REF!", val, "cell.V should be updated even when OnCellCalculated is nil")
	})
}
