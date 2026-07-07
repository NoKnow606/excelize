package excelize

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"path/filepath"
	"strconv"
	"strings"
	"sync"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

func TestGetRangeDataConcurrentLargeSharedStringsStable(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "SharedStrings.xlsx"), Options{UnzipXMLSizeLimit: 128})
	require.NoError(t, err)
	t.Cleanup(func() { _ = f.Close() })

	_, ok := f.tempFiles.Load(defaultXMLPathSharedStrings)
	require.True(t, ok, "expected sharedStrings.xml to use temp-file mode")

	expectedA1, err := f.GetCellValue("Sheet1", "A1")
	require.NoError(t, err)
	require.NotEmpty(t, expectedA1)
	require.False(t, isSharedStringIndexLeak(expectedA1), "baseline A1 should not leak shared string index")

	const workers = 16
	for round := 0; round < 3; round++ {
		var wg sync.WaitGroup
		errCh := make(chan error, workers)
		for i := 0; i < workers; i++ {
			wg.Add(1)
			go func() {
				defer wg.Done()
				data, err := f.GetRangeDataConcurrent("Sheet1", "A1:L2")
				if err != nil {
					errCh <- err
					return
				}
				if len(data) == 0 || len(data[0]) == 0 {
					errCh <- fmt.Errorf("empty range data")
					return
				}
				for _, cell := range data[0] {
					if isSharedStringIndexLeak(cell.Value) {
						errCh <- fmt.Errorf("cell value leaked shared string index: %q", cell.Value)
						return
					}
				}
			}()
		}
		wg.Wait()
		close(errCh)
		for err := range errCh {
			require.NoError(t, err)
		}

		gotA1, err := f.GetCellValue("Sheet1", "A1")
		require.NoError(t, err)
		assert.Equal(t, expectedA1, gotA1, "serial read after concurrent range read must stay stable")
	}
}

func TestFormulaCacheDoesNotPersistCorruptedSharedStringValues(t *testing.T) {
	path := filepath.Join(t.TempDir(), "formula-shared-strings.xlsx")
	f := NewFile()
	require.NoError(t, f.SetCellStr("Sheet1", "A1", "YSD001-家具"))
	require.NoError(t, f.SetCellFormula("Sheet1", "B1", `IFERROR(LEFT(A1,FIND("-",A1)-1),A1)`))
	for row := 2; row <= 80; row++ {
		require.NoError(t, f.SetCellStr("Sheet1", fmt.Sprintf("A%d", row), fmt.Sprintf("SKU-%03d-家具-%s", row, strings.Repeat("x", 16))))
	}
	require.NoError(t, f.SaveAs(path))
	require.NoError(t, f.Close())

	reopened, err := OpenFile(path, Options{UnzipXMLSizeLimit: 128})
	require.NoError(t, err)
	t.Cleanup(func() { _ = reopened.Close() })
	_, ok := reopened.tempFiles.Load(defaultXMLPathSharedStrings)
	require.True(t, ok, "expected sharedStrings.xml to use temp-file mode")

	const workers = 8
	var wg sync.WaitGroup
	errCh := make(chan error, workers)
	for i := 0; i < workers; i++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			data, err := reopened.GetRangeDataConcurrent("Sheet1", "A1:B2")
			if err != nil {
				errCh <- err
				return
			}
			if got := data[0][0].Value; got != "YSD001-家具" {
				errCh <- fmt.Errorf("A1 = %q", got)
			}
		}()
	}
	wg.Wait()
	close(errCh)
	for err := range errCh {
		require.NoError(t, err)
	}

	values, err := reopened.CalcAndUpdateCellValuesDependencyAware("Sheet1", []string{"B1"}, CalcCellValuesDependencyAwareOptions{})
	require.NoError(t, err)
	require.Equal(t, "YSD001", values["B1"])

	out := filepath.Join(t.TempDir(), "formula-shared-strings-recalculated.xlsx")
	require.NoError(t, reopened.SaveAs(out))
	assert.Equal(t, "YSD001", readSavedCellValue(t, out, "xl/worksheets/sheet1.xml", "B1"))
}

func isSharedStringIndexLeak(value string) bool {
	if value == "" {
		return false
	}
	_, err := strconv.Atoi(value)
	return err == nil
}

func readSavedCellValue(t *testing.T, workbookPath string, worksheetPath string, cellRef string) string {
	t.Helper()
	zr, err := zip.OpenReader(workbookPath)
	require.NoError(t, err)
	defer zr.Close()
	for _, file := range zr.File {
		if file.Name != worksheetPath {
			continue
		}
		rc, err := file.Open()
		require.NoError(t, err)
		defer rc.Close()
		decoder := xml.NewDecoder(rc)
		for {
			token, err := decoder.Token()
			if err != nil {
				break
			}
			start, ok := token.(xml.StartElement)
			if !ok || start.Name.Local != "c" {
				continue
			}
			ref := ""
			for _, attr := range start.Attr {
				if attr.Name.Local == "r" {
					ref = attr.Value
					break
				}
			}
			if ref != cellRef {
				continue
			}
			var cell struct {
				V string `xml:"v"`
			}
			require.NoError(t, decoder.DecodeElement(&cell, &start))
			return cell.V
		}
	}
	require.Failf(t, "cell not found", "%s in %s", cellRef, worksheetPath)
	return ""
}
