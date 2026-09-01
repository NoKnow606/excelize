package excelize

import (
	"strings"
	"testing"
)

// TestFormulaToSQL_Tier1_Simple covers basic aggregate and arithmetic functions.
func TestFormulaToSQL_Tier1_Simple(t *testing.T) {
	tests := []struct {
		name    string
		formula string
		wantErr bool
	}{
		// Simple aggregation
		{"T1-SUM-col", "SUM(A:A)", false},
		{"T1-SUM-range", "SUM(A1:A100)", false},
		{"T1-SUM-row", "SUM(1:1)", false},
		{"T1-COUNT-col", "COUNT(A:A)", false},
		{"T1-COUNTA-col", "COUNTA(A:A)", false},
		{"T1-AVERAGE-col", "AVERAGE(A:A)", false},
		{"T1-MAX-col", "MAX(A:A)", false},
		{"T1-MIN-col", "MIN(A:A)", false},
		{"T1-COUNTBLANK", "COUNTBLANK(A:A)", false},

		// Arithmetic
		{"T1-add", "A1+B1", false},
		{"T1-subtract", "A1-B1", false},
		{"T1-multiply", "A1*B1", false},
		{"T1-divide", "A1/B1", false},
		{"T1-power", "A1^2", false},
		{"T1-negate", "-A1", false},
		{"T1-compound", "A1+B1*C1", false},
		{"T1-parens", "(A1+B1)*C1", false},
		{"T1-percent", "A1*5%", false},
		{"T1-sum-plus-sum", "SUM(A:A)+SUM(B:B)", false},

		// Literals
		{"T1-number", "42", false},
		{"T1-string", "\"hello\"", false},
		{"T1-boolean-true", "TRUE", false},
		{"T1-boolean-false", "FALSE", false},

		// Mixed
		{"T1-number-plus-cell", "100+A1", false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			f := NewFile()
			defer f.Close()
			sql, err := f.ConvertFormulaToSQL(tt.formula, FormulaToSQLOptions{SourceTable: "Sheet1"})
			if (err != nil) != tt.wantErr {
				t.Errorf("ConvertFormulaToSQL(%q) error = %v, wantErr %v", tt.formula, err, tt.wantErr)
				return
			}
			if !tt.wantErr && sql == "" {
				t.Errorf("ConvertFormulaToSQL(%q) returned empty SQL", tt.formula)
				return
			}
			t.Logf("Formula: %s → SQL: %s", tt.formula, sql)
		})
	}
}

// TestFormulaToSQL_Tier2_Conditional covers IF, SUMIF, COUNTIF, text functions.
func TestFormulaToSQL_Tier2_Conditional(t *testing.T) {
	tests := []struct {
		name    string
		formula string
		wantErr bool
	}{
		// Conditional aggregates
		{"T2-SUMIF-simple", "SUMIF(A:A,\">0\")", false},
		{"T2-SUMIF-sum-range", "SUMIF(A:A,\"North\",B:B)", false},
		{"T2-SUMIF-text", "SUMIF(A:A,\"Active\")", false},
		{"T2-COUNTIF-gt", "COUNTIF(A:A,\">0\")", false},
		{"T2-COUNTIF-text", "COUNTIF(A:A,\"Completed\")", false},
		{"T2-COUNTIF-equals", "COUNTIF(A:A,A1)", false},
		{"T2-SUMIFS-multi", "SUMIFS(C:C,A:A,\"North\",B:B,\">100\")", false},
		{"T2-COUNTIFS-multi", "COUNTIFS(A:A,\"North\",B:B,\">100\")", false},
		{"T2-AVERAGEIF", "AVERAGEIF(A:A,\">0\")", false},
		{"T2-MAXIFS", "MAXIFS(C:C,A:A,\"North\",B:B,\">100\")", false},
		{"T2-MINIFS", "MINIFS(C:C,A:A,\">0\",B:B,\"North\")", false},

		// Logical
		{"T2-IF-simple", "IF(A1>100,\"High\",\"Low\")", false},
		{"T2-IF-nested", "IF(A1>80,\"A\",IF(A1>60,\"B\",\"C\"))", false},
		{"T2-IF-with-calc", "IF(A1>0,A1*1.1,A1)", false},
		{"T2-IF-with-AND", "IF(AND(A1>0,B1<100),\"OK\",\"Bad\")", false},
		{"T2-IF-with-OR", "IF(OR(A1=\"North\",A1=\"South\"),\"Valid\",\"Invalid\")", false},
		{"T2-IFERROR", "IFERROR(A1/B1,0)", false},
		{"T2-IFERROR-text", "IFERROR(VLOOKUP(A1,B:C,2,FALSE),\"Not Found\")", false},
		{"T2-AND", "AND(A1>0,B1<100)", false},
		{"T2-OR", "OR(A1=\"North\",B1=\"South\")", false},
		{"T2-NOT", "NOT(A1=0)", false},
		{"T2-IFS", "IFS(A1>=90,\"A\",A1>=80,\"B\",A1>=70,\"C\",TRUE,\"F\")", false},
		{"T2-SWITCH", "SWITCH(A1,1,\"Jan\",2,\"Feb\",3,\"Mar\",\"Unknown\")", false},

		// Text
		{"T2-LEFT", "LEFT(A1,3)", false},
		{"T2-RIGHT", "RIGHT(A1,4)", false},
		{"T2-MID", "MID(A1,2,5)", false},
		{"T2-LEN", "LEN(A1)", false},
		{"T2-UPPER", "UPPER(A1)", false},
		{"T2-LOWER", "LOWER(A1)", false},
		{"T2-TRIM", "TRIM(A1)", false},
		{"T2-CONCAT", "CONCAT(A1,B1)", false},
		{"T2-concat-op", "A1&B1", false},
		{"T2-CONCAT-3", "CONCAT(A1,\" \",B1)", false},
		{"T2-SUBSTITUTE", "SUBSTITUTE(A1,\".\",\"-\")", false},
		{"T2-FIND", "FIND(\"abc\",A1)", false},
		{"T2-SEARCH", "SEARCH(\"abc\",A1)", false},
		{"T2-VALUE", "VALUE(A1)", false},
		{"T2-TEXT", "TEXT(A1,\"0.00%\")", false},
		{"T2-REPT", "REPT(\"*\",A1)", false},
		{"T2-TEXTJOIN", "TEXTJOIN(\", \",TRUE,A1,A2,A3)", false},

		// Math
		{"T2-ABS", "ABS(A1)", false},
		{"T2-ROUND", "ROUND(A1,2)", false},
		{"T2-INT", "INT(A1)", false},
		{"T2-SQRT", "SQRT(A1)", false},
		{"T2-MOD", "MOD(A1,7)", false},
		{"T2-POWER", "POWER(A1,2)", false},
		{"T2-SIGN", "SIGN(A1)", false},
		{"T2-EXP", "EXP(A1)", false},
		{"T2-LN", "LN(A1)", false},
		{"T2-LOG10", "LOG10(A1)", false},
		{"T2-LOG-base", "LOG(A1,2)", false},

		// Date
		{"T2-TODAY", "TODAY()", false},
		{"T2-NOW", "NOW()", false},
		{"T2-DATE", "DATE(2024,1,15)", false},
		{"T2-YEAR", "YEAR(A1)", false},
		{"T2-MONTH", "MONTH(A1)", false},
		{"T2-DAY", "DAY(A1)", false},
		{"T2-DATEDIF", "DATEDIF(A1,B1,\"D\")", false},
		{"T2-EDATE", "EDATE(A1,3)", false},
		{"T2-WEEKDAY", "WEEKDAY(A1)", false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			f := NewFile()
			defer f.Close()
			sql, err := f.ConvertFormulaToSQL(tt.formula, FormulaToSQLOptions{SourceTable: "Sheet1"})
			if (err != nil) != tt.wantErr {
				t.Errorf("ConvertFormulaToSQL(%q) error = %v, wantErr %v", tt.formula, err, tt.wantErr)
				return
			}
			if !tt.wantErr && sql == "" {
				t.Errorf("ConvertFormulaToSQL(%q) returned empty SQL", tt.formula)
				return
			}
			t.Logf("Formula: %s → SQL: %s", tt.formula, sql)
		})
	}
}

// TestFormulaToSQL_Tier3_VLOOKUP covers VLOOKUP and HLOOKUP.
func TestFormulaToSQL_Tier3_VLOOKUP(t *testing.T) {
	tests := []struct {
		name    string
		formula string
		wantErr bool
	}{
		{"T3-VLOOKUP-exact", "VLOOKUP(\"SKU001\",A:B,2,FALSE)", false},
		{"T3-VLOOKUP-approx", "VLOOKUP(\"Active\",A:B,2,TRUE)", false},
		{"T3-VLOOKUP-cell-ref", "VLOOKUP(A1,B:D,3,FALSE)", false},
		{"T3-VLOOKUP-cell-ref-approx", "VLOOKUP(A1,A:C,2,TRUE)", false},
		{"T3-HLOOKUP", "HLOOKUP(\"Jan\",A1:Z3,2,FALSE)", false},
		{"T3-LOOKUP-vector", "LOOKUP(A1,B:B,C:C)", false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			f := NewFile()
			defer f.Close()
			sql, err := f.ConvertFormulaToSQL(tt.formula, FormulaToSQLOptions{SourceTable: "Sheet1"})
			if (err != nil) != tt.wantErr {
				t.Errorf("ConvertFormulaToSQL(%q) error = %v, wantErr %v", tt.formula, err, tt.wantErr)
				return
			}
			if !tt.wantErr && sql == "" {
				t.Errorf("ConvertFormulaToSQL(%q) returned empty SQL", tt.formula)
				return
			}
			t.Logf("Formula: %s → SQL: %s", tt.formula, sql)
		})
	}
}

// TestFormulaToSQL_Tier4_INDEX_MATCH covers INDEX, MATCH, INDEX+MATCH combos.
func TestFormulaToSQL_Tier4_INDEX_MATCH(t *testing.T) {
	tests := []struct {
		name    string
		formula string
		wantErr bool
	}{
		// Basic INDEX and MATCH
		{"T4-INDEX-basic", "INDEX(A:A,5)", false},
		{"T4-INDEX-2d", "INDEX(A:C,2,3)", false},
		{"T4-MATCH-basic", "MATCH(\"SKU001\",A:A,0)", false},
		{"T4-MATCH-approx", "MATCH(100,A:A,1)", false},
		{"T4-MATCH-cell", "MATCH(A1,B:B,0)", false},

		// INDEX+MATCH — the critical case
		{"T4-INDEX-MATCH-basic", "INDEX(B:B,MATCH(\"SKU001\",A:A,0))", false},
		{"T4-INDEX-MATCH-cell", "INDEX(C:C,MATCH(A1,B:B,0))", false},
		{"T4-INDEX-MATCH-2col", "INDEX(D:D,MATCH(A1&B1,C:C&F:F,0))", false},
		{"T4-INDEX-MATCH-dynamic", "INDEX(B:B,MATCH(A1,C:C,0))", false},
		{"T4-INDEX-MATCH-price-lookup", "INDEX(C:C,MATCH(A1,B:B,0))", false},
		{"T4-INDEX-MATCH-status", "INDEX(B:B,MATCH(A1,C:C,0))", false},

		// IFERROR-wrapped INDEX+MATCH
		{"T4-IFERROR-INDEX-MATCH", "IFERROR(INDEX(B:B,MATCH(A1,C:C,0)),\"Not Found\")", false},
		{"T4-IFERROR-INDEX-MATCH-zero", "IFERROR(INDEX(C:C,MATCH(A1,D:D,0)),0)", false},

		// XLOOKUP (modern replacement for INDEX+MATCH)
		{"T4-XLOOKUP-basic", "XLOOKUP(A1,B:B,C:C)", false},
		{"T4-XLOOKUP-with-default", "XLOOKUP(A1,B:B,C:C,\"Not Found\")", false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			f := NewFile()
			defer f.Close()
			sql, err := f.ConvertFormulaToSQL(tt.formula, FormulaToSQLOptions{SourceTable: "Sheet1"})
			if (err != nil) != tt.wantErr {
				t.Errorf("ConvertFormulaToSQL(%q) error = %v, wantErr %v", tt.formula, err, tt.wantErr)
				return
			}
			if !tt.wantErr && sql == "" {
				t.Errorf("ConvertFormulaToSQL(%q) returned empty SQL", tt.formula)
				return
			}
			t.Logf("Formula: %s → SQL: %s", tt.formula, sql)
		})
	}
}

// TestFormulaToSQL_Tier5_Modern covers FILTER, SORT, UNIQUE.
func TestFormulaToSQL_Tier5_Modern(t *testing.T) {
	tests := []struct {
		name    string
		formula string
		wantErr bool
	}{
		{"T5-FILTER-basic", "FILTER(A:C,B:B>0)", false},
		{"T5-FILTER-two-cond", "FILTER(A:C,(B:B=\"North\")*(C:C>100))", false},
		{"T5-FILTER-with-if-empty", "FILTER(A:C,B:B>0,\"No Data\")", false},
		{"T5-SORT-ASC", "SORT(A:C,1,1)", false},
		{"T5-SORT-DESC", "SORT(A:C,2,-1)", false},
		{"T5-UNIQUE", "UNIQUE(A:A)", false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			f := NewFile()
			defer f.Close()
			sql, err := f.ConvertFormulaToSQL(tt.formula, FormulaToSQLOptions{SourceTable: "Sheet1"})
			if (err != nil) != tt.wantErr {
				t.Errorf("ConvertFormulaToSQL(%q) error = %v, wantErr %v", tt.formula, err, tt.wantErr)
				return
			}
			if !tt.wantErr && sql == "" {
				t.Errorf("ConvertFormulaToSQL(%q) returned empty SQL", tt.formula)
				return
			}
			t.Logf("Formula: %s → SQL: %s", tt.formula, sql)
		})
	}
}

// TestFormulaToSQL_Tier6_CrossSheet covers cross-sheet references.
func TestFormulaToSQL_Tier6_CrossSheet(t *testing.T) {
	tests := []struct {
		name    string
		formula string
		wantErr bool
	}{
		{"T6-sheet-cell", "Sheet1!A1+Sheet2!B2", false},
		{"T6-sheet-range", "SUM(Products!A:A)", false},
		{"T6-single-quote-sheet", "'Sales Report'!A1", false},
		{"T6-quoted-sheet-range", "SUM('Order Status'!A:A)", false},
		{"T6-VLOOKUP-cross", "VLOOKUP(A1,Products!A:C,3,FALSE)", false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			f := NewFile()
			defer f.Close()
			sql, err := f.ConvertFormulaToSQL(tt.formula, FormulaToSQLOptions{SourceTable: "Sheet1"})
			if (err != nil) != tt.wantErr {
				t.Errorf("ConvertFormulaToSQL(%q) error = %v, wantErr %v", tt.formula, err, tt.wantErr)
				return
			}
			if !tt.wantErr && sql == "" {
				t.Errorf("ConvertFormulaToSQL(%q) returned empty SQL", tt.formula)
				return
			}
			t.Logf("Formula: %s → SQL: %s", tt.formula, sql)
		})
	}
}

// TestFormulaToSQL_Tier7_Ecommerce covers 8 e-commerce scenarios.
func TestFormulaToSQL_Tier7_Ecommerce(t *testing.T) {
	tests := []struct {
		name    string
		formula string
		wantErr bool
	}{
		// 1. Order status lookup — INDEX+MATCH
		{"T7-Ecom-StatusLookup", "IFERROR(INDEX('Order Status'!B:B,MATCH(A1,'Order Status'!A:A,0)),\"Unknown\")", false},

		// 2. Customer tier — nested IF+SUMIF
		{"T7-Ecom-CustomerTier", "IF(SUMIF(C:C,A1,D:D)>10000,\"VIP\",IF(SUMIF(C:C,A1,D:D)>5000,\"Regular\",\"New\"))", false},

		// 3. SKU price lookup — INDEX+MATCH on Products
		{"T7-Ecom-SKUPrice", "IFERROR(INDEX('Products'!C:C,MATCH(A1,'Products'!A:A,0)),0)", false},

		// 4. Stock alert — COUNTIFS
		{"T7-Ecom-StockAlert", "IF(COUNTIFS(A:A,\"Low\",B:B,\"<10\")>0,\"Reorder\",\"OK\")", false},

		// 5. Monthly sales SUMIFS
		{"T7-Ecom-MonthlySales", "SUMIFS(D:D,A:A,\">=2024-01-01\",A:A,\"<=2024-01-31\")", false},

		// 6. SKU profit margin — double INDEX+MATCH
		{"T7-Ecom-SKUProfit", "IFERROR(INDEX('Products'!D:D,MATCH(A1,'Products'!A:A,0))-INDEX('Products'!C:C,MATCH(A1,'Products'!A:A,0)),0)", false},

		// 7. Logistics days — DATEDIF
		{"T7-Ecom-LogisticsDays", "DATEDIF(A1,B1,\"D\")", false},

		// 8. Return rate — COUNTIF ratio
		{"T7-Ecom-ReturnRate", "IFERROR(COUNTIF(A:A,\"Returned\")/COUNTA(A:A),0)", false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			f := NewFile()
			defer f.Close()
			sql, err := f.ConvertFormulaToSQL(tt.formula, FormulaToSQLOptions{SourceTable: "Sheet1"})
			if (err != nil) != tt.wantErr {
				t.Errorf("ConvertFormulaToSQL(%q) error = %v, wantErr %v", tt.formula, err, tt.wantErr)
				return
			}
			if !tt.wantErr && sql == "" {
				t.Errorf("ConvertFormulaToSQL(%q) returned empty SQL", tt.formula)
				return
			}
			t.Logf("Formula: %s → SQL: %s", tt.formula, sql)
		})
	}
}

// TestFormulaToSQL_UnsupportedFunctions verifies that truly unsupported
// functions return an error rather than silently wrong output.
func TestFormulaToSQL_UnsupportedFunctions(t *testing.T) {
	tests := []struct {
		name    string
		formula string
		wantErr bool
	}{
		{"unsupported-SEQUENCE", "SEQUENCE(10)", true},
		{"unsupported-VSTACK", "VSTACK(A:A,B:B)", true},
		{"unsupported-BYROW", "BYROW(A1:C10,LAMBDA(r,SUM(r)))", true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			f := NewFile()
			defer f.Close()
			_, err := f.ConvertFormulaToSQL(tt.formula, FormulaToSQLOptions{SourceTable: "Sheet1"})
			if (err != nil) != tt.wantErr {
				t.Errorf("ConvertFormulaToSQL(%q) error = %v, wantErr %v", tt.formula, err, tt.wantErr)
			}
		})
	}
}

// TestFormulaParser_Tokenizer tests the tokenizer and parser in isolation.
func TestFormulaParser_Tokenizer(t *testing.T) {
	tests := []struct {
		input string
		want  string // substring that should appear in result
	}{
		{"SUM(A:A)", "SUM"},
		{"A1+B1", "A1"},
		{"INDEX(B:B,MATCH(A1,C:C,0))", "INDEX"},
		{"IFERROR(VLOOKUP(A1,B:C,2,FALSE),\"ERR\")", "IFERROR"},
	}
	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			parser := NewFormulaParser(tt.input)
			ast, err := parser.Parse()
			if err != nil {
				t.Fatalf("parse error: %v", err)
			}
			got := ast.Root.String()
			t.Logf("Input: %s → AST: %s", tt.input, got)
			if !strings.Contains(got, tt.want) {
				t.Errorf("AST %q does not contain %q", got, tt.want)
			}
		})
	}
}

// TestFormulaParser_ParseErrors verifies that malformed formulas return errors.
func TestFormulaParser_ParseErrors(t *testing.T) {
	tests := []struct {
		name    string
		formula string
		wantErr bool
	}{
		{"empty", "", false},        // Empty produces EmptyExpr
		{"unterminated-string", `"hello`, true},
		{"unknown-token", "A1 && B1", true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			parser := NewFormulaParser(tt.formula)
			_, err := parser.Parse()
			if (err != nil) != tt.wantErr {
				t.Errorf("Parse(%q) error = %v, wantErr %v", tt.formula, err, tt.wantErr)
			}
		})
	}
}
