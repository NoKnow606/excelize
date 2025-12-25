package excelize

import (
	"fmt"
	"testing"
)

// BenchmarkBatchVsLoop compares batch update performance with loop update
func BenchmarkBatchVsLoop(t *testing.B) {
	sizes := []int{10, 50, 100, 500}

	for _, size := range sizes {
		// Benchmark: Loop with UpdateCellAndRecalculate
		t.Run(fmt.Sprintf("Loop_Update_%d", size), func(b *testing.B) {
			f := NewFile()
			defer f.Close()

			// Set up formula
			f.SetCellFormula("Sheet1", "B1", fmt.Sprintf("=SUM(A1:A%d)", size))

			// Create calcChain
			f.CalcChain = &xlsxCalcChain{
				C: []xlsxCalcChainC{{R: "B1", I: 1}},
			}

			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				// Loop: Update each cell and recalculate
				for j := 1; j <= size; j++ {
					cell, _ := CoordinatesToCellName(1, j)
					f.SetCellValue("Sheet1", cell, j)
					f.UpdateCellAndRecalculate("Sheet1", cell)
				}
			}
		})

		// Benchmark: Batch with BatchUpdateAndRecalculate
		t.Run(fmt.Sprintf("Batch_Update_%d", size), func(b *testing.B) {
			f := NewFile()
			defer f.Close()

			// Set up formula
			f.SetCellFormula("Sheet1", "B1", fmt.Sprintf("=SUM(A1:A%d)", size))

			// Create calcChain
			f.CalcChain = &xlsxCalcChain{
				C: []xlsxCalcChainC{{R: "B1", I: 1}},
			}

			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				// Batch: Update all cells then recalculate once
				updates := make([]CellUpdate, size)
				for j := 1; j <= size; j++ {
					cell, _ := CoordinatesToCellName(1, j)
					updates[j-1] = CellUpdate{
						Sheet: "Sheet1",
						Cell:  cell,
						Value: j,
					}
				}
				f.BatchUpdateAndRecalculate(updates)
			}
		})
	}
}

// BenchmarkBatchSetCellValue compares BatchSetCellValue with loop SetCellValue
func BenchmarkBatchSetCellValue(t *testing.B) {
	sizes := []int{10, 50, 100, 500}

	for _, size := range sizes {
		t.Run(fmt.Sprintf("Loop_SetCellValue_%d", size), func(b *testing.B) {
			f := NewFile()
			defer f.Close()

			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				for j := 1; j <= size; j++ {
					cell, _ := CoordinatesToCellName(1, j)
					f.SetCellValue("Sheet1", cell, j)
				}
			}
		})

		t.Run(fmt.Sprintf("Batch_SetCellValue_%d", size), func(b *testing.B) {
			f := NewFile()
			defer f.Close()

			updates := make([]CellUpdate, size)
			for j := 1; j <= size; j++ {
				cell, _ := CoordinatesToCellName(1, j)
				updates[j-1] = CellUpdate{
					Sheet: "Sheet1",
					Cell:  cell,
					Value: j,
				}
			}

			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				f.BatchSetCellValue(updates)
			}
		})
	}
}

// BenchmarkRecalculateSheet benchmarks RecalculateSheet performance
func BenchmarkRecalculateSheet(t *testing.B) {
	sizes := []int{10, 50, 100, 500}

	for _, size := range sizes {
		t.Run(fmt.Sprintf("Formulas_%d", size), func(b *testing.B) {
			f := NewFile()
			defer f.Close()

			// Set up many formulas
			f.SetCellValue("Sheet1", "A1", 100)
			var calcChainCells []xlsxCalcChainC

			for i := 1; i <= size; i++ {
				cell, _ := CoordinatesToCellName(2, i)
				formula := fmt.Sprintf("=A1*%d", i)
				f.SetCellFormula("Sheet1", cell, formula)

				if i == 1 {
					calcChainCells = append(calcChainCells, xlsxCalcChainC{R: cell, I: 1})
				} else {
					calcChainCells = append(calcChainCells, xlsxCalcChainC{R: cell, I: 0})
				}
			}

			f.CalcChain = &xlsxCalcChain{C: calcChainCells}

			// Pre-calculate once
			f.RecalculateSheet("Sheet1")

			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				f.SetCellValue("Sheet1", "A1", i)
				f.RecalculateSheet("Sheet1")
			}
		})
	}
}

// BenchmarkBatchUpdateMultiSheet benchmarks batch update across multiple sheets
func BenchmarkBatchUpdateMultiSheet(t *testing.B) {
	sheetCounts := []int{2, 5, 10}
	updatesPerSheet := 10

	for _, sheetCount := range sheetCounts {
		t.Run(fmt.Sprintf("Sheets_%d", sheetCount), func(b *testing.B) {
			f := NewFile()
			defer f.Close()

			// Create sheets and set up formulas
			for s := 1; s <= sheetCount; s++ {
				sheetName := fmt.Sprintf("Sheet%d", s)
				if s > 1 {
					f.NewSheet(sheetName)
				}

				f.SetCellFormula(sheetName, "B1", fmt.Sprintf("=SUM(A1:A%d)", updatesPerSheet))
			}

			// Create calcChain for all sheets
			var calcChainCells []xlsxCalcChainC
			for s := 1; s <= sheetCount; s++ {
				calcChainCells = append(calcChainCells, xlsxCalcChainC{R: "B1", I: s})
			}
			f.CalcChain = &xlsxCalcChain{C: calcChainCells}

			// Prepare updates
			updates := make([]CellUpdate, sheetCount*updatesPerSheet)
			idx := 0
			for s := 1; s <= sheetCount; s++ {
				sheetName := fmt.Sprintf("Sheet%d", s)
				for j := 1; j <= updatesPerSheet; j++ {
					cell, _ := CoordinatesToCellName(1, j)
					updates[idx] = CellUpdate{
						Sheet: sheetName,
						Cell:  cell,
						Value: j * s,
					}
					idx++
				}
			}

			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				f.BatchUpdateAndRecalculate(updates)
			}
		})
	}
}

// BenchmarkBatchUpdateComplexFormulas benchmarks with complex formula dependencies
func BenchmarkBatchUpdateComplexFormulas(t *testing.B) {
	t.Run("SimpleSum", func(b *testing.B) {
		f := NewFile()
		defer f.Close()

		f.SetCellFormula("Sheet1", "B1", "=SUM(A1:A100)")
		f.CalcChain = &xlsxCalcChain{C: []xlsxCalcChainC{{R: "B1", I: 1}}}

		updates := make([]CellUpdate, 100)
		for i := 0; i < 100; i++ {
			cell, _ := CoordinatesToCellName(1, i+1)
			updates[i] = CellUpdate{Sheet: "Sheet1", Cell: cell, Value: i + 1}
		}

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			f.BatchUpdateAndRecalculate(updates)
		}
	})

	t.Run("MultipleAggregates", func(b *testing.B) {
		f := NewFile()
		defer f.Close()

		// Set up multiple aggregate functions
		f.SetCellFormula("Sheet1", "B1", "=SUM(A1:A100)")
		f.SetCellFormula("Sheet1", "B2", "=AVERAGE(A1:A100)")
		f.SetCellFormula("Sheet1", "B3", "=MAX(A1:A100)")
		f.SetCellFormula("Sheet1", "B4", "=MIN(A1:A100)")

		f.CalcChain = &xlsxCalcChain{
			C: []xlsxCalcChainC{
				{R: "B1", I: 1},
				{R: "B2", I: 0},
				{R: "B3", I: 0},
				{R: "B4", I: 0},
			},
		}

		updates := make([]CellUpdate, 100)
		for i := 0; i < 100; i++ {
			cell, _ := CoordinatesToCellName(1, i+1)
			updates[i] = CellUpdate{Sheet: "Sheet1", Cell: cell, Value: i + 1}
		}

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			f.BatchUpdateAndRecalculate(updates)
		}
	})
}
