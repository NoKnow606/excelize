package excelize

import (
	"fmt"
	"testing"
)

func TestBatchCalculateINDEXMATCHPatterns(t *testing.T) {
	f := NewFile()
	t.Cleanup(func() { _ = f.Close() })

	if err := f.SetSheetName("Sheet1", "Index1D"); err != nil {
		t.Fatalf("rename default sheet: %v", err)
	}

	sheets := []string{"Lookup1D", "Lookup2D", "LookupAvg", "Index2D", "IndexAvg"}
	for _, sheet := range sheets {
		if _, err := f.NewSheet(sheet); err != nil {
			t.Fatalf("create sheet %s: %v", sheet, err)
		}
	}

	mustSet := func(sheet, cell string, value interface{}) {
		t.Helper()
		if err := f.SetCellValue(sheet, cell, value); err != nil {
			t.Fatalf("set %s!%s failed: %v", sheet, cell, err)
		}
	}
	mustFormula := func(sheet, cell, formula string) {
		t.Helper()
		if err := f.SetCellFormula(sheet, cell, formula); err != nil {
			t.Fatalf("set %s!%s formula failed: %v", sheet, cell, err)
		}
	}

	skuList := []string{"SKU-A", "SKU-B", "SKU-C", "SKU-D", "SKU-E"}
	metrics := []string{"script", "image", "shot"}

	priceMap := make(map[string]float64)
	for i, sku := range skuList {
		row := i + 2
		price := float64(10 * (i + 1))
		mustSet("Lookup1D", fmt.Sprintf("A%d", row), sku)
		mustSet("Lookup1D", fmt.Sprintf("B%d", row), price)
		priceMap[sku] = price
	}
	mustSet("Lookup1D", "A1", "SKU")
	mustSet("Lookup1D", "B1", "Price")

	valueMap := make(map[string]map[string]string)
	mustSet("Lookup2D", "A1", "SKU")
	for j, metric := range metrics {
		colName, _ := ColumnNumberToName(j + 2)
		mustSet("Lookup2D", fmt.Sprintf("%s1", colName), metric)
	}
	for i, sku := range skuList {
		row := i + 2
		mustSet("Lookup2D", fmt.Sprintf("A%d", row), sku)
		valueMap[sku] = make(map[string]string)
		for j, metric := range metrics {
			colName, _ := ColumnNumberToName(j + 2)
			val := fmt.Sprintf("%s_%s", sku, metric)
			mustSet("Lookup2D", fmt.Sprintf("%s%d", colName, row), val)
			valueMap[sku][metric] = val
		}
	}

	avgExpected := make(map[string]float64)
	for i, sku := range skuList {
		row := i + 2
		mustSet("LookupAvg", fmt.Sprintf("A%d", row), sku)
		sum := 0.0
		count := 0
		for col := 0; col < 5; col++ {
			val := float64((i + 1) * (col + 2))
			colName, _ := ColumnNumberToName(col + 3) // start at column C
			mustSet("LookupAvg", fmt.Sprintf("%s%d", colName, row), val)
			sum += val
			count++
		}
		if count > 0 {
			avgExpected[sku] = sum / float64(count)
		}
	}
	mustSet("LookupAvg", "A1", "SKU")

	formulas := make(map[string]string)

	// 1D INDEX-MATCH sheet
	mustSet("Index1D", "A1", "SKU")
	mustSet("Index1D", "B1", "Price")
	for i := 0; i < 12; i++ {
		row := i + 2
		sku := skuList[i%len(skuList)]
		mustSet("Index1D", fmt.Sprintf("A%d", row), sku)
		formula := fmt.Sprintf("=INDEX(Lookup1D!$B:$B,MATCH($A%d,Lookup1D!$A:$A,0))", row)
		cell := fmt.Sprintf("B%d", row)
		mustFormula("Index1D", cell, formula)
		formulas["Index1D!"+cell] = formula
	}

	// 2D INDEX-MATCH sheet
	mustSet("Index2D", "A1", "SKU/Metric")
	for j, metric := range metrics {
		colName, _ := ColumnNumberToName(j + 2)
		mustSet("Index2D", fmt.Sprintf("%s1", colName), metric)
	}
	for i, sku := range skuList[:4] {
		row := i + 2
		mustSet("Index2D", fmt.Sprintf("A%d", row), sku)
		for j := range metrics {
			colName, _ := ColumnNumberToName(j + 2)
			formula := fmt.Sprintf("=INDEX(Lookup2D!$B:$D,MATCH($A%d,Lookup2D!$A:$A,0),MATCH(%s$1,Lookup2D!$B$1:$D$1,0))", row, colName)
			cell := fmt.Sprintf("%s%d", colName, row)
			mustFormula("Index2D", cell, formula)
			formulas["Index2D!"+cell] = formula
		}
	}

	// AVERAGE+INDEX-MATCH sheet
	mustSet("IndexAvg", "A1", "SKU")
	mustSet("IndexAvg", "B1", "Score")
	for i := 0; i < 10; i++ {
		row := i + 2
		sku := skuList[i%len(skuList)]
		mustSet("IndexAvg", fmt.Sprintf("A%d", row), sku)
		formula := fmt.Sprintf("=AVERAGE(INDEX(LookupAvg!$C:$G,MATCH($A%d,LookupAvg!$A:$A,0),0))", row)
		cell := fmt.Sprintf("B%d", row)
		mustFormula("IndexAvg", cell, formula)
		formulas["IndexAvg!"+cell] = formula
	}

	results := f.batchCalculateINDEXMATCH(formulas)

	// Validate 1D result
	targetSku := skuList[0]
	expectedPrice := fmt.Sprintf("%g", priceMap[targetSku])
	if got := results["Index1D!B2"]; got != expectedPrice {
		t.Fatalf("unexpected 1D INDEX-MATCH value, got %s want %s", got, expectedPrice)
	}

	// Validate 2D result
	expected2D := valueMap[skuList[1]][metrics[2]]
	if got := results["Index2D!D3"]; got != expected2D {
		t.Fatalf("unexpected 2D INDEX-MATCH value, got %s want %s", got, expected2D)
	}

	// Validate AVERAGE(INDEX()) result
	expectedAvg := fmt.Sprintf("%g", avgExpected[skuList[2]])
	if got := results["IndexAvg!B4"]; got != expectedAvg {
		t.Fatalf("unexpected AVERAGE(INDEX()) value, got %s want %s", got, expectedAvg)
	}
}

func TestBatchCalculateINDEXMATCHWithCache(t *testing.T) {
	f := NewFile()
	t.Cleanup(func() { _ = f.Close() })

	if err := f.SetSheetName("Sheet1", "CacheIndex1D"); err != nil {
		t.Fatalf("rename sheet: %v", err)
	}
	dataSheets := []string{"CacheLookup", "CacheTableStr", "CacheTableNum", "CacheIndex2D", "CacheIndex2DExpr"}
	for _, sheet := range dataSheets {
		if _, err := f.NewSheet(sheet); err != nil {
			t.Fatalf("create sheet %s: %v", sheet, err)
		}
	}

	mustSet := func(sheet, cell string, value interface{}) {
		t.Helper()
		if err := f.SetCellValue(sheet, cell, value); err != nil {
			t.Fatalf("set %s!%s failed: %v", sheet, cell, err)
		}
	}
	mustFormula := func(sheet, cell, formula string) {
		t.Helper()
		if err := f.SetCellFormula(sheet, cell, formula); err != nil {
			t.Fatalf("set %s!%s formula failed: %v", sheet, cell, err)
		}
	}

	keys := []string{"Key0", "Key1", "Key2", "Key3", "Key4"}
	valueMap := make(map[string]string)
	for i, key := range keys {
		row := i + 2
		val := float64((i + 2) * 3)
		mustSet("CacheLookup", fmt.Sprintf("A%d", row), key)
		mustSet("CacheLookup", fmt.Sprintf("B%d", row), val)
		valueMap[key] = fmt.Sprintf("%g", val)
	}

	metrics := []string{"COL1", "COL2", "COL3"}
	strTable := make(map[string]map[string]string)
	mustSet("CacheTableStr", "A1", "Key")
	for j, metric := range metrics {
		colName, _ := ColumnNumberToName(j + 2)
		mustSet("CacheTableStr", fmt.Sprintf("%s1", colName), metric)
	}
	for i, key := range keys {
		row := i + 2
		mustSet("CacheTableStr", fmt.Sprintf("A%d", row), key)
		strTable[key] = make(map[string]string)
		for j, metric := range metrics {
			colName, _ := ColumnNumberToName(j + 2)
			val := fmt.Sprintf("%s-%s", key, metric)
			mustSet("CacheTableStr", fmt.Sprintf("%s%d", colName, row), val)
			strTable[key][metric] = val
		}
	}

	numMetrics := []int{1, 2, 3}
	numTable := make(map[string]map[int]string)
	mustSet("CacheTableNum", "A1", "Key")
	for j, metric := range numMetrics {
		colName, _ := ColumnNumberToName(j + 2)
		mustSet("CacheTableNum", fmt.Sprintf("%s1", colName), metric)
	}
	for i, key := range keys {
		row := i + 2
		mustSet("CacheTableNum", fmt.Sprintf("A%d", row), key)
		numTable[key] = make(map[int]string)
		for j, metric := range numMetrics {
			colName, _ := ColumnNumberToName(j + 2)
			val := fmt.Sprintf("%d", (i+1)*10+metric)
			mustSet("CacheTableNum", fmt.Sprintf("%s%d", colName, row), val)
			numTable[key][metric] = val
		}
	}

	formulas := make(map[string]string)

	// 1D formulas
	mustSet("CacheIndex1D", "A1", "Key")
	mustSet("CacheIndex1D", "B1", "Value")
	for i := 0; i < 10; i++ {
		row := i + 2
		key := keys[i%len(keys)]
		mustSet("CacheIndex1D", fmt.Sprintf("A%d", row), key)
		formula := fmt.Sprintf("=INDEX(CacheLookup!$B:$B,MATCH($A%d,CacheLookup!$A:$A,0))", row)
		cell := fmt.Sprintf("B%d", row)
		mustFormula("CacheIndex1D", cell, formula)
		formulas["CacheIndex1D!"+cell] = formula
	}

	// 2D string formulas
	mustSet("CacheIndex2D", "A1", "Key")
	for j, metric := range metrics {
		colName, _ := ColumnNumberToName(j + 2)
		mustSet("CacheIndex2D", fmt.Sprintf("%s1", colName), metric)
	}
	for i := 0; i < 4; i++ {
		row := i + 2
		key := keys[i]
		mustSet("CacheIndex2D", fmt.Sprintf("A%d", row), key)
		for j := range metrics {
			colName, _ := ColumnNumberToName(j + 2)
			formula := fmt.Sprintf("=INDEX(CacheTableStr!$B:$D,MATCH($A%d,CacheTableStr!$A:$A,0),MATCH(%s$1,CacheTableStr!$B$1:$D$1,0))", row, colName)
			cell := fmt.Sprintf("%s%d", colName, row)
			mustFormula("CacheIndex2D", cell, formula)
			formulas["CacheIndex2D!"+cell] = formula
		}
	}

	// 2D numeric formulas with +/- adjustments
	mustSet("CacheIndex2DExpr", "A1", "Key")
	for j, metric := range numMetrics {
		colName, _ := ColumnNumberToName(j + 2)
		mustSet("CacheIndex2DExpr", fmt.Sprintf("%s1", colName), metric+1)
	}
	for i := 0; i < 4; i++ {
		row := i + 2
		key := keys[i]
		mustSet("CacheIndex2DExpr", fmt.Sprintf("A%d", row), key)
		for j := range numMetrics {
			colName, _ := ColumnNumberToName(j + 2)
			formula := fmt.Sprintf("=INDEX(CacheTableNum!$B:$D,MATCH($A%d,CacheTableNum!$A:$A,0),MATCH(%s$1-1,CacheTableNum!$B$1:$D$1,0))", row, colName)
			cell := fmt.Sprintf("%s%d", colName, row)
			mustFormula("CacheIndex2DExpr", cell, formula)
			formulas["CacheIndex2DExpr!"+cell] = formula
		}
	}

	cache := NewWorksheetCache()
	for _, sheet := range []string{"CacheLookup", "CacheTableStr", "CacheTableNum", "CacheIndex1D", "CacheIndex2D", "CacheIndex2DExpr"} {
		if err := cache.LoadSheet(f, sheet); err != nil {
			t.Fatalf("load sheet %s: %v", sheet, err)
		}
	}

	results := f.batchCalculateINDEXMATCHWithCache(formulas, cache)

	if got := results["CacheIndex1D!B2"]; got != valueMap[keys[0]] {
		t.Fatalf("unexpected cached 1D result %s", got)
	}
	if got := results["CacheIndex2D!C3"]; got != strTable[keys[1]][metrics[1]] {
		t.Fatalf("unexpected cached 2D result %s", got)
	}
	if got := results["CacheIndex2DExpr!D3"]; got != numTable[keys[1]][numMetrics[2]] {
		t.Fatalf("unexpected +/- adjusted 2D result %s", got)
	}
}

func TestBatchCalculateINDEXMATCHArgsWithCachePreservesTypes(t *testing.T) {
	f := NewFile()
	t.Cleanup(func() { _ = f.Close() })

	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet(): %v", err)
	}

	if err := f.SetCellStr("Data", "A2", "NUM"); err != nil {
		t.Fatalf("SetCellStr Data!A2: %v", err)
	}
	if err := f.SetCellStr("Data", "A3", "STR"); err != nil {
		t.Fatalf("SetCellStr Data!A3: %v", err)
	}
	if err := f.SetCellInt("Data", "B2", 200); err != nil {
		t.Fatalf("SetCellInt Data!B2: %v", err)
	}
	if err := f.SetCellStr("Data", "B3", "200"); err != nil {
		t.Fatalf("SetCellStr Data!B3: %v", err)
	}
	if err := f.SetCellStr("Sheet1", "A1", "NUM"); err != nil {
		t.Fatalf("SetCellStr Sheet1!A1: %v", err)
	}
	if err := f.SetCellStr("Sheet1", "A2", "STR"); err != nil {
		t.Fatalf("SetCellStr Sheet1!A2: %v", err)
	}

	cache := NewWorksheetCache()
	for _, sheet := range []string{"Sheet1", "Data"} {
		if err := cache.LoadSheet(f, sheet); err != nil {
			t.Fatalf("LoadSheet %s: %v", sheet, err)
		}
	}

	results := f.batchCalculateINDEXMATCHArgsWithCache(map[string]string{
		"Sheet1!B1": `=INDEX(Data!$B:$B,MATCH($A1,Data!$A:$A,0))`,
		"Sheet1!B2": `=INDEX(Data!$B:$B,MATCH($A2,Data!$A:$A,0))`,
	}, cache)

	if got := results["Sheet1!B1"]; got.Type != ArgNumber || got.Number != 200 {
		t.Fatalf("expected numeric INDEX-MATCH result, got %#v", got)
	}
	if got := results["Sheet1!B2"]; got.Type != ArgString || got.String != "200" {
		t.Fatalf("expected string INDEX-MATCH result, got %#v", got)
	}
}

func TestBatchIndexMatchRecalculatePreservesTypes(t *testing.T) {
	f := NewFile()
	t.Cleanup(func() { _ = f.Close() })

	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet(): %v", err)
	}

	for i := 1; i <= 5; i++ {
		numRow := i*2 + 0
		strRow := i*2 + 1
		numKey := fmt.Sprintf("NUM%d", i)
		strKey := fmt.Sprintf("STR%d", i)

		if err := f.SetCellStr("Data", fmt.Sprintf("A%d", numRow), numKey); err != nil {
			t.Fatalf("SetCellStr Data!A%d: %v", numRow, err)
		}
		if err := f.SetCellInt("Data", fmt.Sprintf("B%d", numRow), 200); err != nil {
			t.Fatalf("SetCellInt Data!B%d: %v", numRow, err)
		}
		if err := f.SetCellStr("Data", fmt.Sprintf("A%d", strRow), strKey); err != nil {
			t.Fatalf("SetCellStr Data!A%d: %v", strRow, err)
		}
		if err := f.SetCellStr("Data", fmt.Sprintf("B%d", strRow), "200"); err != nil {
			t.Fatalf("SetCellStr Data!B%d: %v", strRow, err)
		}

		targetRow := i*2 - 1
		if err := f.SetCellStr("Sheet1", fmt.Sprintf("A%d", targetRow), numKey); err != nil {
			t.Fatalf("SetCellStr Sheet1!A%d: %v", targetRow, err)
		}
		if err := f.SetCellFormula("Sheet1", fmt.Sprintf("B%d", targetRow), fmt.Sprintf(`IF(TYPE(IFERROR(INDEX(Data!$B:$B,MATCH($A%d,Data!$A:$A,0)),""))=2,"string","number")`, targetRow)); err != nil {
			t.Fatalf("SetCellFormula Sheet1!B%d: %v", targetRow, err)
		}

		targetRow++
		if err := f.SetCellStr("Sheet1", fmt.Sprintf("A%d", targetRow), strKey); err != nil {
			t.Fatalf("SetCellStr Sheet1!A%d: %v", targetRow, err)
		}
		if err := f.SetCellFormula("Sheet1", fmt.Sprintf("B%d", targetRow), fmt.Sprintf(`IF(TYPE(IFERROR(INDEX(Data!$B:$B,MATCH($A%d,Data!$A:$A,0)),""))=2,"string","number")`, targetRow)); err != nil {
			t.Fatalf("SetCellFormula Sheet1!B%d: %v", targetRow, err)
		}
	}

	if err := f.RecalculateAllWithDependency(); err != nil {
		t.Fatalf("RecalculateAllWithDependency(): %v", err)
	}

	for i := 1; i <= 10; i++ {
		want := "number"
		if i%2 == 0 {
			want = "string"
		}
		if got, err := f.GetCellValue("Sheet1", fmt.Sprintf("B%d", i)); err != nil || got != want {
			t.Fatalf("GetCellValue Sheet1!B%d: got %q err=%v want %q", i, got, err, want)
		}
	}
}

func TestDetectAndCalculateBatchINDEX(t *testing.T) {
	f := NewFile()
	t.Cleanup(func() { _ = f.Close() })

	for row := 2; row <= 11; row++ {
		values := []int{row, row + 1, row + 2, row + 3}
		for col, val := range values {
			colName, _ := ColumnNumberToName(col + 3) // start at column C
			if err := f.SetCellValue("Sheet1", fmt.Sprintf("%s%d", colName, row), val); err != nil {
				t.Fatalf("set data: %v", err)
			}
		}
		if err := f.SetCellFormula("Sheet1", fmt.Sprintf("J%d", row), fmt.Sprintf("=INDEX($C%d:$F%d,1,3)", row, row)); err != nil {
			t.Fatalf("set INDEX formula: %v", err)
		}
	}

	results := f.detectAndCalculateBatchINDEX()
	if val, ok := results["Sheet1!J2"]; !ok || int(val) != 4 {
		t.Fatalf("unexpected batch INDEX result %v %v", ok, val)
	}
}
