package excelize

import (
	"fmt"
	"testing"
)

// BenchmarkGetCellStyleNonExistentCells benchmarks reading non-existent cells
// This is where GetCellStyleReadOnly shows significant performance advantage
func BenchmarkGetCellStyleNonExistentCells(b *testing.B) {
	sheet := "Sheet1"

	b.Run("GetCellStyle-NonExistent", func(b *testing.B) {
		f := NewFile()
		defer f.Close()

		// Only create a few cells in column A
		for i := 1; i <= 10; i++ {
			f.SetCellValue(sheet, fmt.Sprintf("A%d", i), i)
		}

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			// Access far-away non-existent cells
			cell := fmt.Sprintf("Z%d", (i%1000)+1)
			f.GetCellStyle(sheet, cell)
		}
	})

	b.Run("GetCellStyleReadOnly-NonExistent", func(b *testing.B) {
		f := NewFile()
		defer f.Close()

		// Only create a few cells in column A
		for i := 1; i <= 10; i++ {
			f.SetCellValue(sheet, fmt.Sprintf("A%d", i), i)
		}

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			// Access far-away non-existent cells
			cell := fmt.Sprintf("Z%d", (i%1000)+1)
			f.GetCellStyleReadOnly(sheet, cell)
		}
	})
}

// BenchmarkGetCellStyleExistingCells benchmarks reading existing cells
// Both methods should have similar performance here
func BenchmarkGetCellStyleExistingCells(b *testing.B) {
	sheet := "Sheet1"

	b.Run("GetCellStyle-Existing", func(b *testing.B) {
		f := NewFile()
		defer f.Close()

		// Create cells with styles
		for i := 1; i <= 100; i++ {
			cell := fmt.Sprintf("A%d", i)
			f.SetCellValue(sheet, cell, i)
		}

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			cell := fmt.Sprintf("A%d", (i%100)+1)
			f.GetCellStyle(sheet, cell)
		}
	})

	b.Run("GetCellStyleReadOnly-Existing", func(b *testing.B) {
		f := NewFile()
		defer f.Close()

		// Create cells with styles
		for i := 1; i <= 100; i++ {
			cell := fmt.Sprintf("A%d", i)
			f.SetCellValue(sheet, cell, i)
		}

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			cell := fmt.Sprintf("A%d", (i%100)+1)
			f.GetCellStyleReadOnly(sheet, cell)
		}
	})
}

// BenchmarkGetCellStyleMemoryImpact benchmarks memory allocation
func BenchmarkGetCellStyleMemoryImpact(b *testing.B) {
	sheet := "Sheet1"

	b.Run("GetCellStyle-MemoryGrowth", func(b *testing.B) {
		f := NewFile()
		defer f.Close()

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			// Access progressively farther cells
			cell := fmt.Sprintf("A%d", i+1)
			f.GetCellStyle(sheet, cell)
		}
		b.StopTimer()

		ws, _ := f.workSheetReader(sheet)
		b.ReportMetric(float64(len(ws.SheetData.Row)), "rows_created")
	})

	b.Run("GetCellStyleReadOnly-MemoryGrowth", func(b *testing.B) {
		f := NewFile()
		defer f.Close()

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			// Access progressively farther cells
			cell := fmt.Sprintf("A%d", i+1)
			f.GetCellStyleReadOnly(sheet, cell)
		}
		b.StopTimer()

		ws, _ := f.workSheetReader(sheet)
		b.ReportMetric(float64(len(ws.SheetData.Row)), "rows_created")
	})
}
