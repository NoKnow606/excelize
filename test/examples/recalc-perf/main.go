package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"runtime"
	"runtime/pprof"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

func main() {
	if len(os.Args) < 2 {
		log.Fatal("Usage: go run ./test/examples/recalc-perf -- <xlsx_file>")
	}

	path := os.Args[1]
	log.Printf("========== Performance Test: %s ==========\n", path)

	// Start CPU profiling
	cpuProfile, err := os.Create("cpu.prof")
	if err != nil {
		log.Printf("Warning: Could not create CPU profile: %v", err)
	} else {
		defer cpuProfile.Close()
		if err := pprof.StartCPUProfile(cpuProfile); err != nil {
			log.Printf("Warning: Could not start CPU profile: %v", err)
		} else {
			defer pprof.StopCPUProfile()
			log.Println("✓ CPU profiling enabled (output: cpu.prof)")
		}
	}

	// Memory baseline
	runtime.GC()
	var memBefore runtime.MemStats
	runtime.ReadMemStats(&memBefore)
	log.Printf("Memory before: %.2f MB (Alloc), %.2f MB (Sys)\n",
		float64(memBefore.Alloc)/1024/1024,
		float64(memBefore.Sys)/1024/1024)

	// Open file
	startOpen := time.Now()
	f, err := excelize.OpenFile(path)
	if err != nil {
		log.Fatalf("Failed to open file: %v", err)
	}
	defer f.Close()
	openDuration := time.Since(startOpen)
	log.Printf("✓ File opened in: %v\n", openDuration)

	// Memory after open
	runtime.GC()
	var memAfterOpen runtime.MemStats
	runtime.ReadMemStats(&memAfterOpen)
	log.Printf("Memory after open: %.2f MB (Alloc), %.2f MB (Sys)\n",
		float64(memAfterOpen.Alloc)/1024/1024,
		float64(memAfterOpen.Sys)/1024/1024)
	log.Printf("Memory increase from open: %.2f MB\n",
		float64(memAfterOpen.Alloc-memBefore.Alloc)/1024/1024)

	// Recalculate with timing
	log.Println("\n--- Starting Recalculation ---")
	startRecalc := time.Now()
	if err := f.RecalculateAllWithDependency(); err != nil {
		log.Fatalf("Recalculation failed: %v", err)
	}
	recalcDuration := time.Since(startRecalc)
	log.Printf("✓ Recalculation completed in: %v\n", recalcDuration)

	// Memory after recalculation
	runtime.GC()
	var memAfterRecalc runtime.MemStats
	runtime.ReadMemStats(&memAfterRecalc)
	log.Printf("Memory after recalc: %.2f MB (Alloc), %.2f MB (Sys)\n",
		float64(memAfterRecalc.Alloc)/1024/1024,
		float64(memAfterRecalc.Sys)/1024/1024)
	log.Printf("Memory increase from recalc: %.2f MB\n",
		float64(memAfterRecalc.Alloc-memAfterOpen.Alloc)/1024/1024)

	// Peak memory
	log.Printf("Peak memory allocated: %.2f MB\n",
		float64(memAfterRecalc.TotalAlloc)/1024/1024)

	// Save result file
	log.Println("\n--- Saving Result File ---")
	// Generate result filename: {origin_file_name}_result.xlsx
	baseFileName := filepath.Base(path)
	ext := filepath.Ext(baseFileName)
	nameWithoutExt := strings.TrimSuffix(baseFileName, ext)
	resultPath := filepath.Join(filepath.Dir(path), nameWithoutExt+"_result.xlsx")

	startSave := time.Now()
	if err := f.SaveAs(resultPath); err != nil {
		log.Fatalf("Failed to save result file: %v", err)
	}
	saveDuration := time.Since(startSave)
	log.Printf("✓ Result saved to: %s\n", resultPath)
	log.Printf("✓ Save completed in: %v\n", saveDuration)

	// Memory after save
	runtime.GC()
	var memAfterSave runtime.MemStats
	runtime.ReadMemStats(&memAfterSave)
	log.Printf("Memory after save: %.2f MB (Alloc), %.2f MB (Sys)\n",
		float64(memAfterSave.Alloc)/1024/1024,
		float64(memAfterSave.Sys)/1024/1024)

	// Summary
	fmt.Println("\n========== Summary ==========")
	fmt.Printf("File: %s\n", path)
	fmt.Printf("Result: %s\n", resultPath)
	fmt.Printf("Open time: %v\n", openDuration)
	fmt.Printf("Recalc time: %v\n", recalcDuration)
	fmt.Printf("Save time: %v\n", saveDuration)
	fmt.Printf("Total time: %v\n", openDuration+recalcDuration+saveDuration)
	fmt.Printf("\nMemory Usage:\n")
	fmt.Printf("  Baseline: %.2f MB\n", float64(memBefore.Alloc)/1024/1024)
	fmt.Printf("  After open: %.2f MB (+%.2f MB)\n",
		float64(memAfterOpen.Alloc)/1024/1024,
		float64(memAfterOpen.Alloc-memBefore.Alloc)/1024/1024)
	fmt.Printf("  After recalc: %.2f MB (+%.2f MB)\n",
		float64(memAfterRecalc.Alloc)/1024/1024,
		float64(memAfterRecalc.Alloc-memAfterOpen.Alloc)/1024/1024)
	fmt.Printf("  After save: %.2f MB (+%.2f MB)\n",
		float64(memAfterSave.Alloc)/1024/1024,
		float64(memAfterSave.Alloc-memAfterRecalc.Alloc)/1024/1024)
	fmt.Printf("  Peak allocated: %.2f MB\n", float64(memAfterSave.TotalAlloc)/1024/1024)
	fmt.Printf("  System memory: %.2f MB\n", float64(memAfterSave.Sys)/1024/1024)
	fmt.Printf("\nGC Stats:\n")
	fmt.Printf("  GC runs: %d\n", memAfterSave.NumGC-memBefore.NumGC)
	fmt.Printf("  GC pause total: %v\n", time.Duration(memAfterSave.PauseTotalNs-memBefore.PauseTotalNs))

	// Write memory profile
	memProfile, err := os.Create("mem.prof")
	if err == nil {
		defer memProfile.Close()
		runtime.GC() // get up-to-date statistics
		if err := pprof.WriteHeapProfile(memProfile); err != nil {
			log.Printf("Warning: Could not write memory profile: %v", err)
		} else {
			log.Println("\n✓ Memory profile saved to: mem.prof")
			log.Println("  Analyze with: go tool pprof mem.prof")
		}
	}

	log.Println("\n✓ Performance test completed!")
}
