package excelize

import (
	"database/sql"
	"fmt"
	"testing"
	"time"

	_ "modernc.org/sqlite"
)

// BenchmarkSQLTransform benchmarks SQL GROUP BY transform for synthetic data
func BenchmarkSQLTransform(b *testing.B) {
	// Define benchmark configurations: rows, driver
	benchmarks := []struct {
		rows   int
		driver string
	}{
		{100000, "sqlite"},
		{500000, "sqlite"},
		{1000000, "sqlite"},
	}

	for _, bb := range benchmarks {
		b.Run(fmt.Sprintf("rows=%d_driver=%s", bb.rows, bb.driver), func(b *testing.B) {
			// Reset timer for benchmark measurement
			b.ResetTimer()
			b.StopTimer()

			for i := 0; i < b.N; i++ {
				// Create in-memory database
				db, err := sql.Open("sqlite", ":memory:")
				if err != nil {
					b.Fatalf("failed to open database: %v", err)
				}
				db.SetMaxOpenConns(1)

				// Create table
				_, err = db.Exec(`CREATE TABLE data (
					id INTEGER PRIMARY KEY,
					category TEXT NOT NULL,
					value REAL NOT NULL,
					timestamp TEXT NOT NULL
				)`)
				if err != nil {
					db.Close()
					b.Fatalf("failed to create table: %v", err)
				}

				// Insert synthetic data
				b.StartTimer()
				insertStart := time.Now()

				tx, err := db.Begin()
				if err != nil {
					b.Fatalf("failed to begin transaction: %v", err)
				}

				stmt, err := tx.Prepare("INSERT INTO data (id, category, value, timestamp) VALUES (?, ?, ?, ?)")
				if err != nil {
					tx.Rollback()
					b.Fatalf("failed to prepare statement: %v", err)
				}

				categories := []string{"A", "B", "C", "D", "E"}
				for j := 0; j < bb.rows; j++ {
					category := categories[j%len(categories)]
					value := float64(j) * 1.5
					timestamp := time.Now().Format(time.RFC3339)
					_, err = stmt.Exec(j, category, value, timestamp)
					if err != nil {
						stmt.Close()
						tx.Rollback()
						b.Fatalf("failed to insert row: %v", err)
					}
				}

				stmt.Close()
				if err = tx.Commit(); err != nil {
					b.Fatalf("failed to commit transaction: %v", err)
				}

				insertDuration := time.Since(insertStart)

				// Execute GROUP BY query
				queryStart := time.Now()
				rows, err := db.Query("SELECT category, SUM(value) FROM data GROUP BY category")
				if err != nil {
					db.Close()
					b.Fatalf("failed to execute query: %v", err)
				}

				resultCount := 0
				for rows.Next() {
					var category string
					var sumValue float64
					if err := rows.Scan(&category, &sumValue); err != nil {
						rows.Close()
						db.Close()
						b.Fatalf("failed to scan row: %v", err)
					}
					resultCount++
				}
				rows.Close()

				if err := rows.Err(); err != nil {
					db.Close()
					b.Fatalf("rows error: %v", err)
				}

				queryDuration := time.Since(queryStart)
				totalDuration := time.Since(insertStart)

				b.StopTimer()

				// Report memory stats
				var memStats struct {
					insertMs  float64
					queryMs   float64
					totalMs   float64
					memoryMB  float64
				}
				memStats.insertMs = float64(insertDuration.Microseconds()) / 1000
				memStats.queryMs = float64(queryDuration.Microseconds()) / 1000
				memStats.totalMs = float64(totalDuration.Microseconds()) / 1000

				// Estimate memory usage (rough approximation)
				// SQLite in-memory: roughly 200 bytes per row + data size
				estimatedMemory := float64(bb.rows) * 200 / 1024 / 1024
				memStats.memoryMB = estimatedMemory

				b.ReportMetric(memStats.insertMs, "insert_ms")
				b.ReportMetric(memStats.queryMs, "query_ms")
				b.ReportMetric(memStats.totalMs, "total_ms")
				b.ReportMetric(memStats.memoryMB, "memory_mb")
				b.ReportMetric(float64(resultCount), "result_count")

				// Log results
				fmt.Printf("rows=%d driver=%s insert_ms=%.2f query_ms=%.2f total_ms=%.2f memory_mb=%.2f\n",
					bb.rows, bb.driver, memStats.insertMs, memStats.queryMs, memStats.totalMs, memStats.memoryMB)

				db.Close()
			}
		})
	}
}
