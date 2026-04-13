# Manual Testing Suite

This directory captures high-level validation assets for Excel compatibility testing that complement automated Go tests.

## Directory Structure

```
test/
├── manual/                          # Manual test plans and workbooks
│   ├── excel2016_system_test_plan.md   # Comprehensive Excel 2016 test plan
│   ├── data/                        # Test workbooks referenced by plans
│   │   ├── test_data_numeric.xlsx
│   │   ├── test_data_text.xlsx
│   │   ├── test_data_date.xlsx
│   │   ├── test_data_mixed.xlsx
│   │   ├── test_data_large.xlsx
│   │   ├── test_data_business.xlsx
│   │   └── excel_formula_parity.xlsx
│   ├── execution_log_template.csv   # PASS/FAIL tracker template
│   ├── execution_log_2024q1.csv     # Test results from Q1 2024
│   └── tools/
│       └── generate_data.go         # Rebuilds test workbooks
├── examples/                        # Sample programs
│   └── recalc.go                    # Recalculation example
├── images/                          # Image files for picture tests
└── *.xlsx                           # Test fixture files

```

## Manual Test Plan

### Excel 2016 Compatibility Tests

The `manual/excel2016_system_test_plan.md` provides a comprehensive test plan originally authored in Chinese for QA teams. It covers:

1. **Basic Formulas** (100+ test cases)
   - Math & Statistics: SUM, AVERAGE, MAX, MIN, COUNT, etc.
   - Logic: IF, AND, OR, IFERROR, etc.
   - Text: CONCATENATE, LEFT, RIGHT, MID, etc.
   - Lookup: VLOOKUP, INDEX, MATCH, etc.
   - Date/Time: DATE, NOW, DATEDIF, etc.

2. **Compound Formulas**
   - Nested IF statements (up to 64 levels)
   - INDEX+MATCH combinations
   - SUMIFS/COUNTIFS with multiple criteria
   - Array formulas

3. **Table Scenarios**
   - Data validation
   - Conditional formatting
   - Pivot tables
   - Named ranges

4. **Boundary Conditions**
   - Large datasets (100k+ rows)
   - Extreme numeric values
   - Empty cells and error values (#DIV/0!, #N/A, etc.)
   - Unicode and special characters

5. **Business Scenarios**
   - Sales analysis with SUMIFS/VLOOKUP
   - Inventory management with complex lookups
   - Financial calculations with nested formulas
   - Attendance tracking with date functions

## How to Use

### 1. Manual Testing Workflow

```bash
# Step 1: Regenerate test data workbooks (if needed)
go run test/manual/tools/generate_data.go

# Step 2: Open the test plan
open test/manual/excel2016_system_test_plan.md

# Step 3: Copy execution log template for this test run
cp test/manual/execution_log_template.csv test/manual/execution_log_$(date +%Y%m%d).csv

# Step 4: Execute test cases and log PASS/FAIL in the CSV
# - Open each test workbook in test/manual/data/
# - Verify formulas calculate correctly
# - Compare results with Excel 2016
# - Mark PASS/FAIL in execution log

# Step 5: Commit execution log to track test history
git add test/manual/execution_log_$(date +%Y%m%d).csv
git commit -m "test: manual testing results for $(date +%Y-%m-%d)"
```

### 2. Test Data Workbooks

All workbooks in `manual/data/` are deterministically generated via `generate_data.go`:

| Workbook | Purpose | Sheets | Notes |
|----------|---------|--------|-------|
| `test_data_numeric.xlsx` | Numeric edge cases | Numbers | Positive/negative integers, decimals, zero, extreme values |
| `test_data_text.xlsx` | Text scenarios | Text | Chinese/English labels, numeric strings, special chars (*, ?, ~, &) |
| `test_data_date.xlsx` | Date handling | Dates | Standard dates, leap years, boundary values, time-of-day |
| `test_data_mixed.xlsx` | Mixed types | Mixed | Blocks of numbers, text, dates, blanks, error values |
| `test_data_large.xlsx` | Large datasets | Values, Text, Lookup, Formulas | 100k+ rows for performance testing |
| `test_data_business.xlsx` | Real-world scenarios | Sales, Inventory, Finance, Attendance, Projects | Business logic validation |

### 3. Rebuilding Test Data

When test requirements change:

```bash
# Edit the generator script
vim test/manual/tools/generate_data.go

# Regenerate all workbooks
go run test/manual/tools/generate_data.go

# Verify changes
git diff test/manual/data/*.xlsx
```

## Execution Logs

### Template Structure

The `execution_log_template.csv` tracks:
- Test Case ID (TC-001, TC-002, etc.)
- Formula being tested
- Expected result
- Actual result
- PASS/FAIL status
- Notes (for failures or edge cases)
- Tester name
- Date

### Historical Results

- `execution_log_2024q1.csv` - First populated run log

Add new files per release to track regression over time.

## Examples

### Recalculation Example (`examples/recalc/main.go`)

Demonstrates how to use the dependency-aware recalculation API:

```go
package main

import (
    "log"
    "github.com/xuri/excelize/v2"
)

func main() {
    f, err := excelize.OpenFile("path/to/workbook.xlsx")
    if err != nil {
        log.Fatal(err)
    }
    defer f.Close()

    // Recalculate all formulas with dependency awareness
    if err := f.RecalculateAllWithDependency(); err != nil {
        log.Fatal(err)
    }

    log.Println("Recalculation completed successfully")
}
```

Run it:
```bash
go run ./test/examples/recalc
```

## Naming Rules

- Lowercase with underscores for workbook names: `test_data_numeric.xlsx`
- Version files by suffixing `_v2` only when schema changes
- Execution logs follow pattern: `execution_log_YYYY[Qq].csv`

## Checklist for New Datasets

Before adding new test workbooks:

- [ ] Workbook name follows naming convention
- [ ] Sheet tab names match test plan references
- [ ] Generator script updated in `tools/generate_data.go`
- [ ] Test plan document updated with new test cases
- [ ] Execution log template updated if new columns needed
- [ ] File metadata includes generation date and maintainer (Workbook → Info → Properties)
- [ ] Workbook committed to version control

## Future Work

- Track PASS/FAIL statistics per release with automated reporting
- Add `test/automation/` directory for automated Excel formula parity checks
- Integrate manual test results into CI pipeline (warning on regressions)
- Support Excel 365 feature testing (dynamic arrays, XLOOKUP, etc.)

---

**Note**: This manual testing suite complements the 100+ automated Go test files. Use automated tests for regression prevention and manual tests for Excel compatibility verification and edge case exploration.
