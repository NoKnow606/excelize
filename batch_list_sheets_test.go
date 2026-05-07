package excelize

import (
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestListSheets(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "TestListSheets.xlsx"))
	assert.NoError(t, err)
	if err != nil {
		return
	}
	defer f.Close()

	assert.Equal(t, []string{"Sheet1", "Sheet2", "Sheet3"}, f.GetSheetList())
}
