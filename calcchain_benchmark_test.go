package excelize

import (
	"fmt"
	"testing"
)

// BenchmarkUpdateCellAndRecalculate benchmarks the performance of UpdateCellAndRecalculate
func BenchmarkUpdateCellAndRecalculate(b *testing.B) {
	sizes := []int{10, 50, 100, 500}

	for _, size := range sizes {
		b.Run(fmt.Sprintf("Chain_%d", size), func(b *testing.B) {
			// Setup: Create a chain of formulas A1 -> A2 -> A3 -> ... -> A{size}
			f := NewFile()
			defer f.Close()

			// Set up the chain
			f.SetCellValue("Sheet1", "A1", 1)
			var calcChainCells []xlsxCalcChainC
			for i := 2; i <= size; i++ {
				cell, _ := CoordinatesToCellName(1, i)
				prevCell, _ := CoordinatesToCellName(1, i-1)
				f.SetCellFormula("Sheet1", cell, "="+prevCell+"+1")
				if i == 2 {
					calcChainCells = append(calcChainCells, xlsxCalcChainC{R: cell, I: 1})
				} else {
					calcChainCells = append(calcChainCells, xlsxCalcChainC{R: cell, I: 0})
				}
			}
			f.CalcChain = &xlsxCalcChain{C: calcChainCells}

			// Pre-calculate to populate cache
			f.UpdateCellAndRecalculate("Sheet1", "A1")

			// Benchmark: Update A1 and recalculate entire chain
			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				f.SetCellValue("Sheet1", "A1", i)
				f.UpdateCellAndRecalculate("Sheet1", "A1")
			}
		})
	}
}

// BenchmarkUpdateCellAndRecalculateWideFormulas benchmarks with many independent formulas
func BenchmarkUpdateCellAndRecalculateWideFormulas(b *testing.B) {
	sizes := []int{10, 50, 100, 500}

	for _, size := range sizes {
		b.Run(fmt.Sprintf("Wide_%d", size), func(b *testing.B) {
			// Setup: Create many formulas that all depend on A1: B1=A1*2, B2=A1*3, B3=A1*4, ...
			f := NewFile()
			defer f.Close()

			f.SetCellValue("Sheet1", "A1", 10)
			var calcChainCells []xlsxCalcChainC
			for i := 1; i <= size; i++ {
				cell, _ := CoordinatesToCellName(2, i) // Column B
				formula := fmt.Sprintf("=A1*%d", i+1)
				f.SetCellFormula("Sheet1", cell, formula)
				if i == 1 {
					calcChainCells = append(calcChainCells, xlsxCalcChainC{R: cell, I: 1})
				} else {
					calcChainCells = append(calcChainCells, xlsxCalcChainC{R: cell, I: 0})
				}
			}
			f.CalcChain = &xlsxCalcChain{C: calcChainCells}

			// Pre-calculate to populate cache
			f.UpdateCellAndRecalculate("Sheet1", "A1")

			// Benchmark: Update A1 and recalculate all dependent formulas
			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				f.SetCellValue("Sheet1", "A1", i)
				f.UpdateCellAndRecalculate("Sheet1", "A1")
			}
		})
	}
}

// BenchmarkUpdateCellAndRecalculateVsCalcCellValue compares UpdateCellAndRecalculate with manual CalcCellValue
func BenchmarkUpdateCellAndRecalculateVsCalcCellValue(b *testing.B) {
	size := 100

	b.Run("UpdateCellAndRecalculate", func(b *testing.B) {
		f := NewFile()
		defer f.Close()

		// Setup chain
		f.SetCellValue("Sheet1", "A1", 1)
		var calcChainCells []xlsxCalcChainC
		for i := 2; i <= size; i++ {
			cell, _ := CoordinatesToCellName(1, i)
			prevCell, _ := CoordinatesToCellName(1, i-1)
			f.SetCellFormula("Sheet1", cell, "="+prevCell+"+1")
			if i == 2 {
				calcChainCells = append(calcChainCells, xlsxCalcChainC{R: cell, I: 1})
			} else {
				calcChainCells = append(calcChainCells, xlsxCalcChainC{R: cell, I: 0})
			}
		}
		f.CalcChain = &xlsxCalcChain{C: calcChainCells}
		f.UpdateCellAndRecalculate("Sheet1", "A1")

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			f.SetCellValue("Sheet1", "A1", i)
			f.UpdateCellAndRecalculate("Sheet1", "A1")
		}
	})

	b.Run("ManualCalcCellValue", func(b *testing.B) {
		f := NewFile()
		defer f.Close()

		// Setup chain
		f.SetCellValue("Sheet1", "A1", 1)
		for i := 2; i <= size; i++ {
			cell, _ := CoordinatesToCellName(1, i)
			prevCell, _ := CoordinatesToCellName(1, i-1)
			f.SetCellFormula("Sheet1", cell, "="+prevCell+"+1")
		}

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			f.SetCellValue("Sheet1", "A1", i)
			// Manually calculate each cell
			for j := 2; j <= size; j++ {
				cell, _ := CoordinatesToCellName(1, j)
				f.CalcCellValue("Sheet1", cell)
			}
		}
	})
}

// BenchmarkUpdateCellCacheVsUpdateCellAndRecalculate compares cache clearing vs immediate recalculation
func BenchmarkUpdateCellCacheVsUpdateCellAndRecalculate(b *testing.B) {
	size := 100

	b.Run("UpdateCellCache", func(b *testing.B) {
		f := NewFile()
		defer f.Close()

		// Setup
		f.SetCellValue("Sheet1", "A1", 1)
		var calcChainCells []xlsxCalcChainC
		for i := 2; i <= size; i++ {
			cell, _ := CoordinatesToCellName(1, i)
			prevCell, _ := CoordinatesToCellName(1, i-1)
			f.SetCellFormula("Sheet1", cell, "="+prevCell+"+1")
			if i == 2 {
				calcChainCells = append(calcChainCells, xlsxCalcChainC{R: cell, I: 1})
			} else {
				calcChainCells = append(calcChainCells, xlsxCalcChainC{R: cell, I: 0})
			}
		}
		f.CalcChain = &xlsxCalcChain{C: calcChainCells}
		f.UpdateCellAndRecalculate("Sheet1", "A1")

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			f.SetCellValue("Sheet1", "A1", i)
			f.UpdateCellCache("Sheet1", "A1") // Only clear cache
		}
	})

	b.Run("UpdateCellAndRecalculate", func(b *testing.B) {
		f := NewFile()
		defer f.Close()

		// Setup
		f.SetCellValue("Sheet1", "A1", 1)
		var calcChainCells []xlsxCalcChainC
		for i := 2; i <= size; i++ {
			cell, _ := CoordinatesToCellName(1, i)
			prevCell, _ := CoordinatesToCellName(1, i-1)
			f.SetCellFormula("Sheet1", cell, "="+prevCell+"+1")
			if i == 2 {
				calcChainCells = append(calcChainCells, xlsxCalcChainC{R: cell, I: 1})
			} else {
				calcChainCells = append(calcChainCells, xlsxCalcChainC{R: cell, I: 0})
			}
		}
		f.CalcChain = &xlsxCalcChain{C: calcChainCells}
		f.UpdateCellAndRecalculate("Sheet1", "A1")

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			f.SetCellValue("Sheet1", "A1", i)
			f.UpdateCellAndRecalculate("Sheet1", "A1") // Calculate and update cache
		}
	})
}
