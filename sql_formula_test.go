package excelize

import (
	"errors"
	"fmt"
	"os"
	"strings"
	"testing"
)

const sqlPostgresTestDSNEnv = "SHEETDB_TEST_DSN"

func TestSQLExecutionBackendOverridesBuiltInPostgresPath(t *testing.T) {
	f := NewFile()
	defer f.Close()

	var received string
	f.SetSQLExecutionBackend(SQLExecutionBackendFunc(func(sqlInput string) (*SQLQueryResult, error) {
		received = sqlInput
		return &SQLQueryResult{
			Columns: []string{"source", "value"},
			Matrix: [][]interface{}{
				{"source", "value"},
				{"postgres", float64(42)},
			},
			SourceSheet: "Data",
		}, nil
	}))

	result, err := f.ExecuteSQL(`select * from gid_0`)
	if err != nil {
		t.Fatalf("ExecuteSQL: %v", err)
	}
	if received != `select * from gid_0` {
		t.Fatalf("backend received %q", received)
	}
	if got := result.Matrix[1][0]; got != "postgres" {
		t.Fatalf("expected backend result, got %#v", got)
	}
}

func TestSQLExecutionBackendUnsupportedRequiresPostgresDSN(t *testing.T) {
	t.Setenv(sqlPostgresDSNEnv, "")
	t.Setenv(sqlPostgresTestDSNEnv, "")
	t.Setenv(sqlPostgresFallbackDSNEnv, "")

	f := NewFile()
	defer f.Close()

	f.SetSQLExecutionBackend(SQLExecutionBackendFunc(func(sqlInput string) (*SQLQueryResult, error) {
		return nil, ErrSQLExecutionBackendUnsupported
	}))

	_, err := f.ExecuteSQL(`select * from "Sheet1"`)
	if !errors.Is(err, ErrSQLPostgresDSNRequired) {
		t.Fatalf("expected PostgreSQL DSN required error, got %v", err)
	}
}

func TestSQLFormulaWithGIDResolverSpillsMatrix(t *testing.T) {
	f := NewFile()
	defer f.Close()

	defaultSheet := f.GetSheetName(0)
	if err := f.SetSheetName(defaultSheet, "Sales"); err != nil {
		t.Fatalf("SetSheetName: %v", err)
	}
	if _, err := f.NewSheet("Report"); err != nil {
		t.Fatalf("NewSheet: %v", err)
	}

	writeSQLSheetRows(t, f, "Sales", [][]interface{}{
		{"Region", "Revenue"},
		{"North", 10},
		{"South", 20},
		{"South", 5},
	})

	f.SetSQLSourceResolver(SQLSourceResolverFunc(func(token string, sheetList []string) (string, error) {
		switch unquoteIdentifier(token) {
		case "gid_7":
			return "Sales", nil
		case "Sales":
			return "Sales", nil
		default:
			return "", fmt.Errorf("unexpected token %s", token)
		}
	}))

	formula := `SQL("select ""Region"", sum(""Revenue"") as ""Total"" from gid_7 group by ""Region"" order by ""Total"" desc")`
	if err := f.SetCellFormula("Report", "A1", formula); err != nil {
		t.Fatalf("SetCellFormula: %v", err)
	}

	result, err := f.CalcCellValueWithMatrix("Report", "A1")
	if err != nil {
		t.Fatalf("CalcCellValueWithMatrix: %v", err)
	}
	if len(result.Matrix) != 3 {
		t.Fatalf("expected header plus 2 rows, got %#v", result.Matrix)
	}
	if got := result.Matrix[0][0]; got != "Region" {
		t.Fatalf("expected header Region, got %#v", got)
	}
	if got := result.Matrix[1][0]; got != "South" {
		t.Fatalf("expected first data row South, got %#v", got)
	}
	if got := result.Matrix[1][1]; got != float64(25) {
		t.Fatalf("expected South total 25, got %#v", got)
	}
	if got := result.Matrix[2][0]; got != "North" {
		t.Fatalf("expected second data row North, got %#v", got)
	}
}

func TestSQLFormulaSupportsDerivedTableSubqueries(t *testing.T) {
	f := NewFile()
	defer f.Close()

	defaultSheet := f.GetSheetName(0)
	if err := f.SetSheetName(defaultSheet, "Sales"); err != nil {
		t.Fatalf("SetSheetName: %v", err)
	}
	if _, err := f.NewSheet("Report"); err != nil {
		t.Fatalf("NewSheet: %v", err)
	}

	writeSQLSheetRows(t, f, "Sales", [][]interface{}{
		{"Region", "Revenue"},
		{"North", "10"},
		{"South", "20"},
		{"South", "5"},
	})

	f.SetSQLSourceResolver(SQLSourceResolverFunc(func(token string, sheetList []string) (string, error) {
		switch unquoteIdentifier(token) {
		case "gid_7":
			return "Sales", nil
		default:
			return "", fmt.Errorf("unexpected token %s", token)
		}
	}))

	formula := `SQL("select Region, sum(amount) as ""Total"" from (select ""Region"" as Region, cast(""Revenue"" as integer) as amount from gid_7) as source group by Region order by ""Total"" desc")`
	if err := f.SetCellFormula("Report", "A1", formula); err != nil {
		t.Fatalf("SetCellFormula: %v", err)
	}

	result, err := f.CalcCellValueWithMatrix("Report", "A1")
	if err != nil {
		t.Fatalf("CalcCellValueWithMatrix: %v", err)
	}
	if len(result.Matrix) != 3 {
		t.Fatalf("expected header plus 2 rows, got %#v", result.Matrix)
	}
	if got := result.Matrix[0][0]; got != "Region" && got != "region" {
		t.Fatalf("expected header Region/region, got %#v", got)
	}
	if got := result.Matrix[1][0]; got != "South" {
		t.Fatalf("expected first data row South, got %#v", got)
	}
	if got := result.Matrix[1][1]; got != float64(25) {
		t.Fatalf("expected South total 25, got %#v", got)
	}
	if got := result.Matrix[2][0]; got != "North" {
		t.Fatalf("expected second data row North, got %#v", got)
	}

	sourceSheets := f.sqlFormulaSourceSheets(formula)
	if len(sourceSheets) != 1 || sourceSheets[0] != "Sales" {
		t.Fatalf("expected nested query dependency on Sales, got %#v", sourceSheets)
	}
}

func TestSQLFormulaSupportsProvidedCleanedDataCTE(t *testing.T) {
	f := NewFile()
	defer f.Close()

	defaultSheet := f.GetSheetName(0)
	if err := f.SetSheetName(defaultSheet, "Traffic"); err != nil {
		t.Fatalf("SetSheetName: %v", err)
	}
	if _, err := f.NewSheet("Report"); err != nil {
		t.Fatalf("NewSheet: %v", err)
	}

	writeSQLSheetRows(t, f, "Traffic", [][]interface{}{
		{"店铺", "SPU分类", "主商品货号", "商品访客（访问）", "商品访客（添加至购物车）"},
		{"Shop-A", "服饰", "SKU-1", "10", "2"},
		{"Shop-A", "服饰", "SKU-2", "5", "1"},
		{"Shop-A", "鞋类", "ABC", "3", "0"},
		{"Shop-B", "服饰", "SKU-9", "-", "4"},
	})

	f.SetSQLSourceResolver(SQLSourceResolverFunc(func(token string, sheetList []string) (string, error) {
		switch unquoteIdentifier(token) {
		case "gid_0":
			return "Traffic", nil
		default:
			return "", fmt.Errorf("unexpected token %s", token)
		}
	}))

	formula := `SQL("
WITH cleaned_data AS (
    SELECT 
        LOWER(""店铺"") AS shop,
        ""SPU分类"" AS category,
        CASE
            WHEN INSTR(""主商品货号"", '-') > 0
            THEN SUBSTR(""主商品货号"", 1, INSTR(""主商品货号"", '-') - 1)
            ELSE ""主商品货号""
        END AS main_product_sku,
        CAST(""商品访客（访问）"" AS INTEGER) AS visit_amount,
        CAST(""商品访客（添加至购物车）"" AS INTEGER) AS add_cart_amount
    FROM gid_0
    WHERE ""商品访客（访问）"" <> '-'
)

SELECT 
    shop,
    category,
    main_product_sku AS ""主商品货号"",
    SUM(visit_amount) AS visit_sum,
    SUM(add_cart_amount) AS add_cart_sum,
    shop || '-' || category || '-' || main_product_sku AS KID
FROM cleaned_data
GROUP BY 
    shop, 
    category, 
    main_product_sku
")`
	if err := f.SetCellFormula("Report", "A1", formula); err != nil {
		t.Fatalf("SetCellFormula: %v", err)
	}

	result, err := f.CalcCellValueWithMatrix("Report", "A1")
	if err != nil {
		t.Fatalf("CalcCellValueWithMatrix: %v", err)
	}
	if len(result.Matrix) != 3 {
		t.Fatalf("expected header plus 2 rows, got %#v", result.Matrix)
	}

	rowsByKID := make(map[string][]interface{}, len(result.Matrix)-1)
	for _, row := range result.Matrix[1:] {
		rowsByKID[row[5].(string)] = row
	}

	row := rowsByKID["shop-a-服饰-SKU"]
	if row == nil {
		t.Fatalf("expected grouped row for shop-a-服饰-SKU, got %#v", result.Matrix)
	}
	if row[0] != "shop-a" || row[1] != "服饰" || row[2] != "SKU" || row[3] != float64(15) || row[4] != float64(3) {
		t.Fatalf("unexpected grouped row for shop-a-服饰-SKU: %#v", row)
	}

	row = rowsByKID["shop-a-鞋类-ABC"]
	if row == nil {
		t.Fatalf("expected grouped row for shop-a-鞋类-ABC, got %#v", result.Matrix)
	}
	if row[0] != "shop-a" || row[1] != "鞋类" || row[2] != "ABC" || row[3] != float64(3) || row[4] != float64(0) {
		t.Fatalf("unexpected grouped row for shop-a-鞋类-ABC: %#v", row)
	}

	sourceSheets := f.sqlFormulaSourceSheets(formula)
	if len(sourceSheets) != 1 || sourceSheets[0] != "Traffic" {
		t.Fatalf("expected CTE query dependency on Traffic, got %#v", sourceSheets)
	}
}

func TestSQLFormulaSupportsProvidedBaseDataCTE(t *testing.T) {
	f := NewFile()
	defer f.Close()

	defaultSheet := f.GetSheetName(0)
	if err := f.SetSheetName(defaultSheet, "Spend"); err != nil {
		t.Fatalf("SetSheetName: %v", err)
	}
	if _, err := f.NewSheet("Report"); err != nil {
		t.Fatalf("NewSheet: %v", err)
	}

	writeSQLSheetRows(t, f, "Spend", [][]interface{}{
		{"店铺", "SPU分类", "主货号", "花费"},
		{"马来站", "配件", "AA-1", "100"},
		{"泰国站", "配件", "AA-2", "200"},
		{"CN Shop", "箱包", "BB", "50.5"},
	})

	f.SetSQLSourceResolver(SQLSourceResolverFunc(func(token string, sheetList []string) (string, error) {
		switch unquoteIdentifier(token) {
		case "gid_2":
			return "Spend", nil
		default:
			return "", fmt.Errorf("unexpected token %s", token)
		}
	}))

	formula := `SQL("
WITH base_data AS (
    SELECT 
        LOWER(""店铺"") AS shop,
        ""SPU分类"" AS category,
        CASE
            WHEN INSTR(""主货号"", '-') > 0
            THEN SUBSTR(""主货号"", 1, INSTR(""主货号"", '-') - 1)
            ELSE ""主货号""
        END AS main_sku_clean,
        COALESCE(CAST(""花费"" AS REAL), 0) AS cost_amount,
        LOWER(""店铺"") || '-' || ""SPU分类"" || '-' ||
        CASE
            WHEN INSTR(""主货号"", '-') > 0
            THEN SUBSTR(""主货号"", 1, INSTR(""主货号"", '-') - 1)
            ELSE ""主货号""
        END AS KID
    FROM gid_2
)

SELECT 
    shop,
    category,
    main_sku_clean AS ""主货号"",
    SUM(
        CASE 
            WHEN shop LIKE '%马来%' OR shop LIKE '%malaysia%' OR shop LIKE '%my%' 
            THEN cost_amount * 1.08
            WHEN shop LIKE '%泰国%' OR shop LIKE '%thailand%' OR shop LIKE '%th%' 
            THEN cost_amount * 1.07
            ELSE cost_amount
        END
    ) AS cost_sum,
    KID
FROM base_data
GROUP BY 
    shop, 
    category, 
    main_sku_clean, 
    KID;
")`
	if err := f.SetCellFormula("Report", "A1", formula); err != nil {
		t.Fatalf("SetCellFormula: %v", err)
	}

	result, err := f.CalcCellValueWithMatrix("Report", "A1")
	if err != nil {
		t.Fatalf("CalcCellValueWithMatrix: %v", err)
	}
	if len(result.Matrix) != 4 {
		t.Fatalf("expected header plus 3 rows, got %#v", result.Matrix)
	}

	rowsByKID := make(map[string][]interface{}, len(result.Matrix)-1)
	for _, row := range result.Matrix[1:] {
		rowsByKID[row[4].(string)] = row
	}

	row := rowsByKID["马来站-配件-AA"]
	if row == nil {
		t.Fatalf("expected row for 马来站-配件-AA, got %#v", result.Matrix)
	}
	if row[0] != "马来站" || row[1] != "配件" || row[2] != "AA" || row[3] != float64(108) {
		t.Fatalf("unexpected row for 马来站-配件-AA: %#v", row)
	}

	row = rowsByKID["泰国站-配件-AA"]
	if row == nil {
		t.Fatalf("expected row for 泰国站-配件-AA, got %#v", result.Matrix)
	}
	if row[0] != "泰国站" || row[1] != "配件" || row[2] != "AA" || row[3] != float64(214) {
		t.Fatalf("unexpected row for 泰国站-配件-AA: %#v", row)
	}

	row = rowsByKID["cn shop-箱包-BB"]
	if row == nil {
		t.Fatalf("expected row for cn shop-箱包-BB, got %#v", result.Matrix)
	}
	if row[0] != "cn shop" || row[1] != "箱包" || row[2] != "BB" || row[3] != float64(50.5) {
		t.Fatalf("unexpected row for cn shop-箱包-BB: %#v", row)
	}

	sourceSheets := f.sqlFormulaSourceSheets(formula)
	if len(sourceSheets) != 1 || sourceSheets[0] != "Spend" {
		t.Fatalf("expected CTE query dependency on Spend, got %#v", sourceSheets)
	}
}

func TestSQLFormulaFindAffectedCellsByCellsTracksWholeSourceSheet(t *testing.T) {
	f := NewFile()
	defer f.Close()

	defaultSheet := f.GetSheetName(0)
	if err := f.SetSheetName(defaultSheet, "Data"); err != nil {
		t.Fatalf("SetSheetName: %v", err)
	}
	if _, err := f.NewSheet("Summary"); err != nil {
		t.Fatalf("NewSheet: %v", err)
	}

	writeSQLSheetRows(t, f, "Data", [][]interface{}{
		{"Category", "Amount"},
		{"A", 10},
		{"B", 20},
	})

	formula := `SQL("select ""Category"", sum(""Amount"") as ""Total"" from ""Data"" group by ""Category"" order by ""Category""")`
	if err := f.SetCellFormula("Summary", "A1", formula); err != nil {
		t.Fatalf("SetCellFormula: %v", err)
	}

	graph := f.buildDependencyGraph()
	node := graph.nodes["Summary!A1"]
	if node == nil {
		t.Fatal("expected Summary!A1 in dependency graph")
	}
	if !containsString(node.dependencies, "SHEET:Data") {
		t.Fatalf("expected whole-sheet dependency marker, got %#v", node.dependencies)
	}

	affected := f.findAffectedCellsByCells(graph, map[string]bool{"Data!B2": true})
	if !affected["Summary!A1"] {
		t.Fatalf("expected Summary!A1 to be affected by Data!B2 update, got %#v", affected)
	}
}

func TestSQLFormulaUpdateFormulaCachePersistsSpillRange(t *testing.T) {
	f := NewFile()
	defer f.Close()

	defaultSheet := f.GetSheetName(0)
	if err := f.SetSheetName(defaultSheet, "Data"); err != nil {
		t.Fatalf("SetSheetName: %v", err)
	}
	if _, err := f.NewSheet("Report"); err != nil {
		t.Fatalf("NewSheet: %v", err)
	}

	writeSQLSheetRows(t, f, "Data", [][]interface{}{
		{"Category", "Amount"},
		{"A", 10},
		{"B", 20},
	})

	formula := `SQL("select ""Category"", sum(""Amount"") as ""Total"" from ""Data"" group by ""Category"" order by ""Category""")`
	if err := f.SetCellFormula("Report", "A1", formula); err != nil {
		t.Fatalf("SetCellFormula: %v", err)
	}

	if err := f.UpdateFormulaCache(); err != nil {
		t.Fatalf("UpdateFormulaCache: %v", err)
	}

	ws, err := f.workSheetReader("Report")
	if err != nil {
		t.Fatalf("workSheetReader: %v", err)
	}
	if got := ws.SheetData.Row[0].C[0].F.Ref; got != "A1:B3" {
		t.Fatalf("expected spill ref A1:B3, got %q", got)
	}

	if got, err := f.GetCellValue("Report", "B3", Options{RawCellValue: true}); err != nil || got != "20" {
		t.Fatalf("expected in-memory spill value 20 at B3, got %q err=%v", got, err)
	}

	fileName := "test_sql_formula_cache.xlsx"
	if err := f.SaveAs(fileName); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}
	defer os.Remove(fileName)

	reopened, err := OpenFile(fileName)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer reopened.Close()

	if got, err := reopened.GetCellValue("Report", "A1", Options{RawCellValue: true}); err != nil || got != "Category" {
		t.Fatalf("expected reopened A1 header Category, got %q err=%v", got, err)
	}
	if got, err := reopened.GetCellValue("Report", "B2", Options{RawCellValue: true}); err != nil || got != "10" {
		t.Fatalf("expected reopened B2 value 10, got %q err=%v", got, err)
	}
	if got, err := reopened.GetCellValue("Report", "B3", Options{RawCellValue: true}); err != nil || got != "20" {
		t.Fatalf("expected reopened B3 value 20, got %q err=%v", got, err)
	}
}

func TestSQLFormulaStoreCalculatedValuePersistsSpillRangeAndCaches(t *testing.T) {
	f := NewFile()
	defer f.Close()

	defaultSheet := f.GetSheetName(0)
	if err := f.SetSheetName(defaultSheet, "Data"); err != nil {
		t.Fatalf("SetSheetName: %v", err)
	}
	if _, err := f.NewSheet("Report"); err != nil {
		t.Fatalf("NewSheet: %v", err)
	}

	writeSQLSheetRows(t, f, "Data", [][]interface{}{
		{"Category", "Amount"},
		{"A", 10},
		{"B", 20},
	})

	formula := `SQL("select ""Category"", sum(""Amount"") as ""Total"" from ""Data"" group by ""Category"" order by ""Category""")`
	if err := f.SetCellFormula("Report", "A1", formula); err != nil {
		t.Fatalf("SetCellFormula: %v", err)
	}

	worksheetCache := NewWorksheetCache()
	f.storeCalculatedValue("Report", "A1", "Category", worksheetCache)

	if cached, ok := worksheetCache.Get("Report", "B2"); !ok || cached.Type != ArgNumber || cached.Number != 10 {
		t.Fatalf("expected worksheet cache numeric B2=10, got ok=%v arg=%+v", ok, cached)
	}
	if cached, ok := f.calcCache.Load("Report!B3!raw=true"); !ok || cached.(string) != "20" {
		t.Fatalf("expected calc cache B3 raw value 20, got ok=%v value=%v", ok, cached)
	}

	ws, err := f.workSheetReader("Report")
	if err != nil {
		t.Fatalf("workSheetReader: %v", err)
	}
	if got := ws.SheetData.Row[0].C[0].F.Ref; got != "A1:B3" {
		t.Fatalf("expected spill ref A1:B3, got %q", got)
	}

	if got, err := f.GetCellValue("Report", "B2", Options{RawCellValue: true}); err != nil || got != "10" {
		t.Fatalf("expected in-memory B2 value 10, got %q err=%v", got, err)
	}

	fileName := "test_sql_formula_trigger.xlsx"
	if err := f.SaveAs(fileName); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}
	defer os.Remove(fileName)

	reopened, err := OpenFile(fileName)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer reopened.Close()

	if got, err := reopened.GetCellValue("Report", "A1", Options{RawCellValue: true}); err != nil || got != "Category" {
		t.Fatalf("expected reopened A1 header Category, got %q err=%v", got, err)
	}
	if got, err := reopened.GetCellValue("Report", "B3", Options{RawCellValue: true}); err != nil || got != "20" {
		t.Fatalf("expected reopened B3 value 20, got %q err=%v", got, err)
	}
}

func TestSQLFormulaBatchUpdatePersistsSpillRangeAndCaches(t *testing.T) {
	f := NewFile()
	defer f.Close()

	defaultSheet := f.GetSheetName(0)
	if err := f.SetSheetName(defaultSheet, "Data"); err != nil {
		t.Fatalf("SetSheetName: %v", err)
	}
	if _, err := f.NewSheet("Report"); err != nil {
		t.Fatalf("NewSheet: %v", err)
	}

	writeSQLSheetRows(t, f, "Data", [][]interface{}{
		{"Category", "Amount"},
		{"A", 10},
		{"B", 20},
	})

	formula := `SQL("select ""Category"", sum(""Amount"") as ""Total"" from ""Data"" group by ""Category"" order by ""Category""")`
	if err := f.BatchUpdateValuesAndFormulasWithRecalcV2(
		[]CellUpdate{{Sheet: "Report", Cell: "A1", Value: "Category"}},
		[]FormulaUpdateWithValue{{Sheet: "Report", Cell: "A1", Formula: formula}},
	); err != nil {
		t.Fatalf("BatchUpdateValuesAndFormulasWithRecalcV2: %v", err)
	}

	ws, err := f.workSheetReader("Report")
	if err != nil {
		t.Fatalf("workSheetReader: %v", err)
	}
	if got := ws.SheetData.Row[0].C[0].F.Ref; got != "A1:B3" {
		t.Fatalf("expected spill ref A1:B3, got %q", got)
	}
	if got, err := f.GetCellValue("Report", "B3", Options{RawCellValue: true}); err != nil || got != "20" {
		t.Fatalf("expected in-memory spill value 20 at B3, got %q err=%v", got, err)
	}
	if cached, ok := f.calcCache.Load("Report!B2!raw=true"); !ok || cached.(string) != "10" {
		t.Fatalf("expected calc cache B2 raw value 10, got ok=%v value=%v", ok, cached)
	}

	fileName := "test_sql_formula_batch_update.xlsx"
	if err := f.SaveAs(fileName); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}
	defer os.Remove(fileName)

	reopened, err := OpenFile(fileName)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer reopened.Close()

	if got, err := reopened.GetCellValue("Report", "A1", Options{RawCellValue: true}); err != nil || got != "Category" {
		t.Fatalf("expected reopened A1 header Category, got %q err=%v", got, err)
	}
	if got, err := reopened.GetCellValue("Report", "B2", Options{RawCellValue: true}); err != nil || got != "10" {
		t.Fatalf("expected reopened B2 value 10, got %q err=%v", got, err)
	}
	if got, err := reopened.GetCellValue("Report", "B3", Options{RawCellValue: true}); err != nil || got != "20" {
		t.Fatalf("expected reopened B3 value 20, got %q err=%v", got, err)
	}
}

func TestSQLFormulaClearsPreviousSpillOnError(t *testing.T) {
	f := NewFile()
	defer f.Close()

	defaultSheet := f.GetSheetName(0)
	if err := f.SetSheetName(defaultSheet, "Data"); err != nil {
		t.Fatalf("SetSheetName: %v", err)
	}
	if _, err := f.NewSheet("Report"); err != nil {
		t.Fatalf("NewSheet: %v", err)
	}

	writeSQLSheetRows(t, f, "Data", [][]interface{}{
		{"Category", "Amount"},
		{"A", 10},
		{"B", 20},
	})

	validFormula := `SQL("select ""Category"", sum(""Amount"") as ""Total"" from ""Data"" group by ""Category"" order by ""Category""")`
	if err := f.SetCellFormula("Report", "A1", validFormula); err != nil {
		t.Fatalf("SetCellFormula: %v", err)
	}

	worksheetCache := NewWorksheetCache()
	f.storeCalculatedValue("Report", "A1", "Category", worksheetCache)

	if _, ok := f.calcCache.Load("Report!B3!raw=true"); !ok {
		t.Fatal("expected spill cache for Report!B3 before error")
	}

	invalidFormula := `SQL("delete from ""Data""")`
	if err := f.SetCellFormula("Report", "A1", invalidFormula); err != nil {
		t.Fatalf("SetCellFormula invalid SQL: %v", err)
	}
	f.setFormulaValue("Report", "A1", formulaErrorVALUE)

	ws, err := f.workSheetReader("Report")
	if err != nil {
		t.Fatalf("workSheetReader: %v", err)
	}
	if got := ws.SheetData.Row[0].C[0].F.Ref; got != "" {
		t.Fatalf("expected spill ref cleared after error, got %q", got)
	}

	if got, err := f.GetCellValue("Report", "A1", Options{RawCellValue: true}); err != nil || got != formulaErrorVALUE {
		t.Fatalf("expected A1 error value %s, got %q err=%v", formulaErrorVALUE, got, err)
	}
	if got, err := f.GetCellValue("Report", "B2", Options{RawCellValue: true}); err != nil || got != "" {
		t.Fatalf("expected cleared spill cell B2, got %q err=%v", got, err)
	}
	if got, err := f.GetCellValue("Report", "B3", Options{RawCellValue: true}); err != nil || got != "" {
		t.Fatalf("expected cleared spill cell B3, got %q err=%v", got, err)
	}
	if _, ok := f.calcCache.Load("Report!B3!raw=true"); ok {
		t.Fatal("expected spill cache for Report!B3 to be cleared after error")
	}
}

func TestSQLFormulaShrinksPreviousSpillAfterFormulaChange(t *testing.T) {
	f := NewFile()
	defer f.Close()

	defaultSheet := f.GetSheetName(0)
	if err := f.SetSheetName(defaultSheet, "Data"); err != nil {
		t.Fatalf("SetSheetName: %v", err)
	}
	if _, err := f.NewSheet("Report"); err != nil {
		t.Fatalf("NewSheet: %v", err)
	}

	writeSQLSheetRows(t, f, "Data", [][]interface{}{
		{"Category", "Amount"},
		{"A", 10},
		{"B", 20},
	})

	limitTwo := `SQL("select ""Category"", ""Amount"" from ""Data"" order by ""Category"" limit 2")`
	if err := f.SetCellFormula("Report", "A1", limitTwo); err != nil {
		t.Fatalf("SetCellFormula limitTwo: %v", err)
	}
	if err := f.RecalculateAllWithDependency(); err != nil {
		t.Fatalf("RecalculateAllWithDependency limitTwo: %v", err)
	}

	ws, err := f.workSheetReader("Report")
	if err != nil {
		t.Fatalf("workSheetReader: %v", err)
	}
	if got := ws.SheetData.Row[0].C[0].F.Ref; got != "A1:B3" {
		t.Fatalf("expected initial spill ref A1:B3, got %q", got)
	}
	if got, err := f.GetCellValue("Report", "A3", Options{RawCellValue: true}); err != nil || got != "B" {
		t.Fatalf("expected initial A3 spill value B, got %q err=%v", got, err)
	}
	if got, err := f.GetCellValue("Report", "B3", Options{RawCellValue: true}); err != nil || got != "20" {
		t.Fatalf("expected initial B3 spill value 20, got %q err=%v", got, err)
	}

	limitOne := `SQL("select ""Category"", ""Amount"" from ""Data"" order by ""Category"" limit 1")`
	if err := f.SetCellFormula("Report", "A1", limitOne); err != nil {
		t.Fatalf("SetCellFormula limitOne: %v", err)
	}
	if err := f.RecalculateAllWithDependency(); err != nil {
		t.Fatalf("RecalculateAllWithDependency limitOne: %v", err)
	}

	ws, err = f.workSheetReader("Report")
	if err != nil {
		t.Fatalf("workSheetReader after shrink: %v", err)
	}
	if got := ws.SheetData.Row[0].C[0].F.Ref; got != "A1:B2" {
		t.Fatalf("expected shrunk spill ref A1:B2, got %q", got)
	}
	if got, err := f.GetCellValue("Report", "A3", Options{RawCellValue: true}); err != nil || got != "" {
		t.Fatalf("expected cleared spill cell A3, got %q err=%v", got, err)
	}
	if got, err := f.GetCellValue("Report", "B3", Options{RawCellValue: true}); err != nil || got != "" {
		t.Fatalf("expected cleared spill cell B3, got %q err=%v", got, err)
	}

	fileName := "test_sql_formula_shrink.xlsx"
	if err := f.SaveAs(fileName); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}
	defer os.Remove(fileName)

	reopened, err := OpenFile(fileName)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer reopened.Close()

	if got, err := reopened.GetCellValue("Report", "A3", Options{RawCellValue: true}); err != nil || got != "" {
		t.Fatalf("expected reopened cleared spill cell A3, got %q err=%v", got, err)
	}
	if got, err := reopened.GetCellValue("Report", "B3", Options{RawCellValue: true}); err != nil || got != "" {
		t.Fatalf("expected reopened cleared spill cell B3, got %q err=%v", got, err)
	}
}

func TestSetCellFormulaWithValueClearsPreviousSQLSpillImmediately(t *testing.T) {
	f := NewFile()
	defer f.Close()

	defaultSheet := f.GetSheetName(0)
	if err := f.SetSheetName(defaultSheet, "Data"); err != nil {
		t.Fatalf("SetSheetName: %v", err)
	}
	if _, err := f.NewSheet("Report"); err != nil {
		t.Fatalf("NewSheet: %v", err)
	}

	writeSQLSheetRows(t, f, "Data", [][]interface{}{
		{"Category", "Amount"},
		{"A", 10},
		{"B", 20},
	})

	limitTwo := `SQL("select ""Category"", ""Amount"" from ""Data"" order by ""Category"" limit 2")`
	if err := f.SetCellFormula("Report", "A1", limitTwo); err != nil {
		t.Fatalf("SetCellFormula limitTwo: %v", err)
	}
	if err := f.RecalculateAllWithDependency(); err != nil {
		t.Fatalf("RecalculateAllWithDependency limitTwo: %v", err)
	}

	limitOne := `SQL("select ""Category"", ""Amount"" from ""Data"" order by ""Category"" limit 1")`
	if err := f.SetCellFormulaWithValue("Report", "A1", limitOne, "Category"); err != nil {
		t.Fatalf("SetCellFormulaWithValue limitOne: %v", err)
	}

	ws, err := f.workSheetReader("Report")
	if err != nil {
		t.Fatalf("workSheetReader after setter: %v", err)
	}
	if got := ws.SheetData.Row[0].C[0].F.Ref; got != "" {
		t.Fatalf("expected stale spill ref cleared before recalc, got %q", got)
	}
	if got, err := f.GetCellValue("Report", "A3", Options{RawCellValue: true}); err != nil || got != "" {
		t.Fatalf("expected immediate clear of stale spill cell A3, got %q err=%v", got, err)
	}
	if got, err := f.GetCellValue("Report", "B3", Options{RawCellValue: true}); err != nil || got != "" {
		t.Fatalf("expected immediate clear of stale spill cell B3, got %q err=%v", got, err)
	}

	fileName := "test_sql_formula_setter_clear.xlsx"
	if err := f.SaveAs(fileName); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}
	defer os.Remove(fileName)

	reopened, err := OpenFile(fileName)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer reopened.Close()

	if got, err := reopened.GetCellValue("Report", "A3", Options{RawCellValue: true}); err != nil || got != "" {
		t.Fatalf("expected reopened stale spill cell A3 cleared, got %q err=%v", got, err)
	}
	if got, err := reopened.GetCellValue("Report", "B3", Options{RawCellValue: true}); err != nil || got != "" {
		t.Fatalf("expected reopened stale spill cell B3 cleared, got %q err=%v", got, err)
	}
}

func TestSQLFormulaShrinkRefreshesUsedRangeForWideSpill(t *testing.T) {
	f := NewFile()
	defer f.Close()

	defaultSheet := f.GetSheetName(0)
	if err := f.SetSheetName(defaultSheet, "Data"); err != nil {
		t.Fatalf("SetSheetName: %v", err)
	}
	if _, err := f.NewSheet("Report"); err != nil {
		t.Fatalf("NewSheet: %v", err)
	}

	headers := make([]interface{}, 23)
	rowOne := make([]interface{}, 23)
	rowTwo := make([]interface{}, 23)
	for i := 0; i < 23; i++ {
		headers[i] = fmt.Sprintf("col_%02d", i+1)
		rowOne[i] = fmt.Sprintf("r1c%02d", i+1)
		rowTwo[i] = fmt.Sprintf("r2c%02d", i+1)
	}
	writeSQLSheetRows(t, f, "Data", [][]interface{}{
		headers,
		rowOne,
		rowTwo,
	})

	limitTwo := `SQL("select * from ""Data"" limit 2")`
	if err := f.SetCellFormula("Report", "A1", limitTwo); err != nil {
		t.Fatalf("SetCellFormula limitTwo: %v", err)
	}
	if err := f.RecalculateAllWithDependency(); err != nil {
		t.Fatalf("RecalculateAllWithDependency limitTwo: %v", err)
	}
	if got, err := f.GetSheetDimension("Report"); err != nil || got != "A1:W3" {
		t.Fatalf("expected wide spill dimension A1:W3, got %q err=%v", got, err)
	}

	limitOne := `SQL("select * from ""Data"" limit 1")`
	if err := f.SetCellFormulaWithValue("Report", "A1", limitOne, "col_01"); err != nil {
		t.Fatalf("SetCellFormulaWithValue limitOne: %v", err)
	}
	if err := f.RecalculateAllWithDependency(); err != nil {
		t.Fatalf("RecalculateAllWithDependency limitOne: %v", err)
	}
	if got, err := f.GetSheetDimension("Report"); err != nil || got != "A1:W2" {
		t.Fatalf("expected shrunk dimension A1:W2, got %q err=%v", got, err)
	}

	fileName := "test_sql_formula_wide_shrink.xlsx"
	if err := f.SaveAs(fileName); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}
	defer os.Remove(fileName)

	reopened, err := OpenFile(fileName)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer reopened.Close()

	if got, err := reopened.GetSheetDimension("Report"); err != nil || got != "A1:W2" {
		t.Fatalf("expected reopened dimension A1:W2, got %q err=%v", got, err)
	}
	rangeData, err := reopened.GetRangeDataConcurrent("Report", "A1:W2")
	if err != nil {
		t.Fatalf("GetRangeDataConcurrent A1:W2: %v", err)
	}
	if len(rangeData) != 2 || len(rangeData[0]) != 23 {
		t.Fatalf("expected 2x23 data after shrink, got %dx%d", len(rangeData), len(rangeData[0]))
	}
	if got := rangeData[1][22].Value; got != "r1c23" {
		t.Fatalf("expected second row last value r1c23, got %q", got)
	}
	staleRow, err := reopened.GetRangeDataConcurrent("Report", "A3:W3")
	if err != nil {
		t.Fatalf("GetRangeDataConcurrent A3:W3: %v", err)
	}
	for _, cell := range staleRow[0] {
		if cell.Value != "" || cell.Formula != "" {
			t.Fatalf("expected stale third row to be empty, got value=%q formula=%q", cell.Value, cell.Formula)
		}
	}
}

func TestRefreshWorksheetDimensionIgnoresStyleOnlyPlaceholderCells(t *testing.T) {
	f := NewFile()
	defer f.Close()

	defaultSheet := f.GetSheetName(0)
	if err := f.SetSheetName(defaultSheet, "Report"); err != nil {
		t.Fatalf("SetSheetName: %v", err)
	}

	if err := f.SetCellValue("Report", "A1", "header"); err != nil {
		t.Fatalf("SetCellValue A1: %v", err)
	}
	if err := f.SetCellValue("Report", "W2", "tail"); err != nil {
		t.Fatalf("SetCellValue W2: %v", err)
	}

	styleID, err := f.NewStyle(&Style{})
	if err != nil {
		t.Fatalf("NewStyle: %v", err)
	}
	if err := f.SetCellStyle("Report", "Z1000", "Z1000", styleID); err != nil {
		t.Fatalf("SetCellStyle Z1000: %v", err)
	}

	ws, err := f.workSheetReader("Report")
	if err != nil {
		t.Fatalf("workSheetReader: %v", err)
	}
	refreshWorksheetDimension(ws)

	if got, err := f.GetSheetDimension("Report"); err != nil || got != "A1:W2" {
		t.Fatalf("expected dimension A1:W2 ignoring style-only placeholder, got %q err=%v", got, err)
	}
}

func TestClearWorksheetRangeValuesPreservesContiguousCellIndexing(t *testing.T) {
	f := NewFile()
	defer f.Close()

	defaultSheet := f.GetSheetName(0)
	if err := f.SetSheetName(defaultSheet, "Report"); err != nil {
		t.Fatalf("SetSheetName: %v", err)
	}

	ws, err := f.workSheetReader("Report")
	if err != nil {
		t.Fatalf("workSheetReader: %v", err)
	}

	for row := 1; row <= 2; row++ {
		for col := 1; col <= 26; col++ {
			cell, _ := CoordinatesToCellName(col, row)
			c, _, _, err := ws.prepareCell(cell)
			if err != nil {
				t.Fatalf("prepareCell %s: %v", cell, err)
			}
			c.V = fmt.Sprintf("old-r%dc%d", row, col)
			c.T = "str"
		}
	}

	clearWorksheetRangeValues(ws, "A1:Z2", "A1")

	for row := 1; row <= 2; row++ {
		for col := 1; col <= 23; col++ {
			cell, _ := CoordinatesToCellName(col, row)
			c, _, _, err := ws.prepareCell(cell)
			if err != nil {
				t.Fatalf("prepareCell rewrite %s: %v", cell, err)
			}
			c.V = fmt.Sprintf("new-r%dc%d", row, col)
			c.T = "str"
		}
	}
	refreshWorksheetDimension(ws)

	rangeData, err := f.GetRangeDataConcurrent("Report", "A1:Z2")
	if err != nil {
		t.Fatalf("GetRangeDataConcurrent: %v", err)
	}
	if got := rangeData[0][0].Value; got != "new-r1c1" {
		t.Fatalf("expected A1 new-r1c1, got %q", got)
	}
	if got := rangeData[0][22].Value; got != "new-r1c23" {
		t.Fatalf("expected W1 new-r1c23, got %q", got)
	}
	if got := rangeData[1][0].Value; got != "new-r2c1" {
		t.Fatalf("expected A2 new-r2c1, got %q", got)
	}
	if got := rangeData[1][22].Value; got != "new-r2c23" {
		t.Fatalf("expected W2 new-r2c23, got %q", got)
	}
	for col := 23; col < 26; col++ {
		if got := rangeData[0][col].Value; got != "" {
			t.Fatalf("expected cleared tail at row1 col %d, got %q", col+1, got)
		}
		if got := rangeData[1][col].Value; got != "" {
			t.Fatalf("expected cleared tail at row2 col %d, got %q", col+1, got)
		}
	}
}

func TestClearWorksheetRangeValuesDropsDefaultHeightPlaceholderRows(t *testing.T) {
	f := NewFile()
	defer f.Close()

	defaultSheet := f.GetSheetName(0)
	if err := f.SetSheetName(defaultSheet, "Report"); err != nil {
		t.Fatalf("SetSheetName: %v", err)
	}

	ws, err := f.workSheetReader("Report")
	if err != nil {
		t.Fatalf("workSheetReader: %v", err)
	}
	ws.SheetFormatPr = &xlsxSheetFormatPr{
		DefaultRowHeight: 18,
		CustomHeight:     true,
	}

	cells := []string{"A1", "B1", "A3", "B3", "A6", "B6"}
	for _, cell := range cells {
		c, _, _, err := ws.prepareCell(cell)
		if err != nil {
			t.Fatalf("prepareCell %s: %v", cell, err)
		}
		c.V = cell
		c.T = "str"
	}

	clearWorksheetRangeValues(ws, "A1:B6", "A1")

	for _, cell := range []string{"A1", "B1", "A2", "B2"} {
		c, _, _, err := ws.prepareCell(cell)
		if err != nil {
			t.Fatalf("prepareCell rewrite %s: %v", cell, err)
		}
		c.V = "new-" + cell
		c.T = "str"
	}
	refreshWorksheetDimension(ws)

	if len(ws.SheetData.Row) != 2 {
		t.Fatalf("expected only 2 rows after pruning placeholder rows, got %d", len(ws.SheetData.Row))
	}
	if ws.SheetData.Row[0].R != 1 || ws.SheetData.Row[1].R != 2 {
		t.Fatalf("expected rows 1 and 2 after pruning, got %d and %d", ws.SheetData.Row[0].R, ws.SheetData.Row[1].R)
	}
	if got, err := f.GetSheetDimension("Report"); err != nil || got != "A1:B2" {
		t.Fatalf("expected dimension A1:B2, got %q err=%v", got, err)
	}

	rangeData, err := f.GetRangeDataConcurrent("Report", "A1:B6")
	if err != nil {
		t.Fatalf("GetRangeDataConcurrent: %v", err)
	}
	if len(rangeData) != 6 {
		t.Fatalf("expected 6 rows in requested range, got %d", len(rangeData))
	}
	if got := rangeData[0][0].Value; got != "new-A1" {
		t.Fatalf("expected A1 new-A1, got %q", got)
	}
	if got := rangeData[1][1].Value; got != "new-B2" {
		t.Fatalf("expected B2 new-B2, got %q", got)
	}
	for row := 2; row < 6; row++ {
		for col := 0; col < 2; col++ {
			if got := rangeData[row][col].Value; got != "" {
				t.Fatalf("expected cleared placeholder row %d col %d, got %q", row+1, col+1, got)
			}
		}
	}
}

func TestClearWorksheetRangeValuesPreservesInteriorEmptyRowsForLaterWrites(t *testing.T) {
	f := NewFile()
	defer f.Close()

	defaultSheet := f.GetSheetName(0)
	if err := f.SetSheetName(defaultSheet, "Report"); err != nil {
		t.Fatalf("SetSheetName: %v", err)
	}

	ws, err := f.workSheetReader("Report")
	if err != nil {
		t.Fatalf("workSheetReader: %v", err)
	}

	for _, cell := range []string{"A1", "B1", "A3", "B3"} {
		c, _, _, err := ws.prepareCell(cell)
		if err != nil {
			t.Fatalf("prepareCell %s: %v", cell, err)
		}
		c.V = "old-" + cell
		c.T = "str"
	}

	clearWorksheetRangeValues(ws, "A1:B1", "A1")

	c, _, _, err := ws.prepareCell("A2")
	if err != nil {
		t.Fatalf("prepareCell A2: %v", err)
	}
	c.V = "new-A2"
	c.T = "str"
	refreshWorksheetDimension(ws)

	rangeData, err := f.GetRangeDataConcurrent("Report", "A1:B3")
	if err != nil {
		t.Fatalf("GetRangeDataConcurrent: %v", err)
	}
	if got := rangeData[1][0].Value; got != "new-A2" {
		t.Fatalf("expected A2 new-A2, got %q", got)
	}
	if got := rangeData[2][0].Value; got != "old-A3" {
		t.Fatalf("expected A3 old-A3 to stay on row 3, got %q", got)
	}
	if got := rangeData[2][1].Value; got != "old-B3" {
		t.Fatalf("expected B3 old-B3 to stay on row 3, got %q", got)
	}
}

func TestSQLFormulaRecalculateRepairsSparseRowSlice(t *testing.T) {
	f := NewFile()
	defer f.Close()

	defaultSheet := f.GetSheetName(0)
	if err := f.SetSheetName(defaultSheet, "Data"); err != nil {
		t.Fatalf("SetSheetName: %v", err)
	}
	if _, err := f.NewSheet("Report"); err != nil {
		t.Fatalf("NewSheet Report: %v", err)
	}

	writeSQLSheetRows(t, f, "Data", [][]interface{}{
		{"Category", "Amount"},
		{"A", 10},
		{"B", 20},
	})

	limitTwo := `SQL("select ""Category"", ""Amount"" from ""Data"" order by ""Category"" limit 2")`
	if err := f.SetCellFormula("Report", "A1", limitTwo); err != nil {
		t.Fatalf("SetCellFormula limitTwo: %v", err)
	}
	if err := f.RecalculateAllWithDependency(); err != nil {
		t.Fatalf("RecalculateAllWithDependency limitTwo: %v", err)
	}

	ws, err := f.workSheetReader("Report")
	if err != nil {
		t.Fatalf("workSheetReader Report: %v", err)
	}
	if len(ws.SheetData.Row) < 3 {
		t.Fatalf("expected at least 3 rows after initial spill, got %d", len(ws.SheetData.Row))
	}
	ws.SheetData.Row = append(ws.SheetData.Row[:1], ws.SheetData.Row[2:]...)

	limitOne := `SQL("select ""Category"", ""Amount"" from ""Data"" order by ""Category"" limit 1")`
	if err := f.SetCellFormulaWithValue("Report", "A1", limitOne, "Category"); err != nil {
		t.Fatalf("SetCellFormulaWithValue limitOne: %v", err)
	}
	if err := f.RecalculateAllWithDependency(); err != nil {
		t.Fatalf("RecalculateAllWithDependency limitOne: %v", err)
	}

	rangeData, err := f.GetRangeDataConcurrent("Report", "A1:B3")
	if err != nil {
		t.Fatalf("GetRangeDataConcurrent: %v", err)
	}
	if got := rangeData[0][0].Value; got != "Category" {
		t.Fatalf("expected A1 header Category, got %q", got)
	}
	if got := rangeData[1][0].Value; got != "A" {
		t.Fatalf("expected A2 repaired to A, got %q", got)
	}
	if got := rangeData[1][1].Value; got != "10" {
		t.Fatalf("expected B2 repaired to 10, got %q", got)
	}
	if got := rangeData[2][0].Value; got != "" {
		t.Fatalf("expected A3 cleared after shrink, got %q", got)
	}
	if got := rangeData[2][1].Value; got != "" {
		t.Fatalf("expected B3 cleared after shrink, got %q", got)
	}
}

func TestSQLFormulaShrinkClearsStaleCellsBeyondStoredSpillRef(t *testing.T) {
	f := NewFile()
	defer f.Close()

	defaultSheet := f.GetSheetName(0)
	if err := f.SetSheetName(defaultSheet, "Data"); err != nil {
		t.Fatalf("SetSheetName: %v", err)
	}
	if _, err := f.NewSheet("Report"); err != nil {
		t.Fatalf("NewSheet Report: %v", err)
	}

	writeSQLSheetRows(t, f, "Data", [][]interface{}{
		{"Category", "Amount"},
		{"A", 10},
		{"B", 20},
	})

	limitTwo := `SQL("select ""Category"", ""Amount"" from ""Data"" order by ""Category"" limit 2")`
	if err := f.SetCellFormula("Report", "A1", limitTwo); err != nil {
		t.Fatalf("SetCellFormula limitTwo: %v", err)
	}
	if err := f.RecalculateAllWithDependency(); err != nil {
		t.Fatalf("RecalculateAllWithDependency limitTwo: %v", err)
	}

	if err := f.SetCellValue("Report", "A6", "stale-A6"); err != nil {
		t.Fatalf("SetCellValue A6: %v", err)
	}
	if err := f.SetCellValue("Report", "Z3", "stale-Z3"); err != nil {
		t.Fatalf("SetCellValue Z3: %v", err)
	}

	limitOne := `SQL("select ""Category"", ""Amount"" from ""Data"" order by ""Category"" limit 1")`
	if err := f.SetCellFormulaWithValue("Report", "A1", limitOne, "Category"); err != nil {
		t.Fatalf("SetCellFormulaWithValue limitOne: %v", err)
	}
	if err := f.RecalculateAllWithDependency(); err != nil {
		t.Fatalf("RecalculateAllWithDependency limitOne: %v", err)
	}

	if got, err := f.GetCellValue("Report", "A6"); err != nil || got != "" {
		t.Fatalf("expected stale A6 cleared, got %q err=%v", got, err)
	}
	if got, err := f.GetCellValue("Report", "Z3"); err != nil || got != "" {
		t.Fatalf("expected stale Z3 cleared, got %q err=%v", got, err)
	}
	if got, err := f.GetSheetDimension("Report"); err != nil || got != "A1:B2" {
		t.Fatalf("expected shrunk dimension A1:B2, got %q err=%v", got, err)
	}
}

func writeSQLSheetRows(t *testing.T, f *File, sheet string, rows [][]interface{}) {
	t.Helper()
	requirePostgresSQL(t, f)
	for rowIdx, row := range rows {
		cell, err := CoordinatesToCellName(1, rowIdx+1)
		if err != nil {
			t.Fatalf("CoordinatesToCellName: %v", err)
		}
		if err := f.SetSheetRow(sheet, cell, &row); err != nil {
			t.Fatalf("SetSheetRow: %v", err)
		}
	}
}

func requirePostgresSQL(t *testing.T, f *File) {
	t.Helper()
	dsn := strings.TrimSpace(os.Getenv(sqlPostgresDSNEnv))
	if dsn == "" {
		dsn = strings.TrimSpace(os.Getenv(sqlPostgresTestDSNEnv))
	}
	if dsn == "" {
		dsn = strings.TrimSpace(os.Getenv(sqlPostgresFallbackDSNEnv))
	}
	if dsn == "" {
		t.Skipf("PostgreSQL SQL formula test requires %s, %s, or %s", sqlPostgresDSNEnv, sqlPostgresTestDSNEnv, sqlPostgresFallbackDSNEnv)
	}
	f.SetSQLPostgresDSN(dsn)
}

func containsString(values []string, want string) bool {
	for _, value := range values {
		if value == want {
			return true
		}
	}
	return false
}
