# SQL Formula Architecture

This document explains how the native `SQL("...")` formula works in `omnimcp-excelize`, how PostgreSQL is used for query execution, and the main tradeoffs.

## Overview

The SQL formula is a custom workbook-aware formula function. It lets a sheet cell run a read-only SQL query against worksheet data and return the result as a spill matrix.

Example:

```excel
=SQL("select ""Category"", sum(""Amount"") as ""Total"" from ""Data"" group by ""Category"" order by ""Category""")
```

The first row of the query result becomes the output header row. The full result spills from the anchor cell into the worksheet.

## Primary Purpose

The feature exists to solve a gap between spreadsheet formulas and relational data processing.

Use it when you want to:

- aggregate worksheet rows with `GROUP BY`
- filter and reshape data with `WHERE`, `CASE`, `CAST`, and `ORDER BY`
- express report logic more clearly than nested Excel formulas
- generate report tables from one anchor formula cell
- keep the logic inside the workbook formula engine instead of exporting to an external SQL step

In practice, it acts like an embedded analytical query layer over workbook sheets.

## How It Works

At a high level:

1. The formula engine recognizes `SQL(...)`.
2. The SQL text is extracted and validated.
3. Worksheet names or custom tokens such as `gid_7` are resolved.
4. Referenced worksheets are copied into temporary PostgreSQL tables.
5. PostgreSQL executes the query.
6. The result is converted to a matrix and spilled back into the worksheet.
7. The spilled cells and spill range metadata are persisted so the workbook can be saved and reopened with the computed result intact.

### Execution Flow

```mermaid
flowchart TD
    A["Cell formula with SQL query"] --> B["Parse SQL string argument"]
    B --> C["Validate single SELECT or WITH query"]
    C --> D["Find FROM and JOIN source tokens"]
    D --> E["Resolve worksheet names or gid_* tokens"]
    E --> F["Rewrite sources to internal PostgreSQL table names"]
    F --> G["Open PostgreSQL connection from DSN"]
    G --> H["Read worksheet rows via GetRows"]
    H --> I["Build temporary PostgreSQL tables from sheet headers and rows"]
    I --> J["Execute rewritten SQL"]
    J --> K["Convert result rows to formula matrix"]
    K --> L["Spill matrix into worksheet cells"]
    L --> M["Persist spill ref and cached values"]
```

### Sequence Diagram

```mermaid
sequenceDiagram
    participant User as Workbook Cell
    participant Engine as Formula Engine
    participant Resolver as SQL Source Resolver
    participant Postgres as External PostgreSQL
    participant Sheet as Source Worksheet
    participant Target as Target Worksheet

    User->>Engine: Evaluate SQL query
    Engine->>Engine: Extract and validate query
    Engine->>Resolver: Resolve source token
    Resolver-->>Engine: Resolve source token to worksheet name
    Engine->>Postgres: Open DSN connection
    Engine->>Sheet: Read rows from source worksheet
    Sheet-->>Engine: Header row + data rows
    Engine->>Postgres: CREATE TEMP TABLE + INSERT rows
    Engine->>Postgres: Execute rewritten SELECT
    Postgres-->>Engine: Result columns + result rows
    Engine->>Target: Write spill cells and spill ref
    Target-->>User: Anchor cell returns top-left result
```

## Storage Model

There are two different storage layers involved.

### 1. PostgreSQL Execution Storage

The SQL formula engine uses an external PostgreSQL database process via DSN.

- The DSN comes from `File.SetSQLPostgresDSN(...)`, `EXCELIZE_SQL_POSTGRES_DSN`, or `POSTGRES_DSN`.
- Each compile or execute call opens a short-lived PostgreSQL connection.
- Referenced worksheet data is materialized into temporary PostgreSQL tables on that connection.
- Compatibility functions and aggregates are installed in the `excelize_sql_compat` schema and used through `search_path`.
- The connection is closed after compilation or execution finishes.

This means the workbook SQL formula path no longer embeds or opens SQLite. PostgreSQL must be reachable when using the built-in SQL formula execution path.

### 2. Workbook Storage

The query result is persisted into normal worksheet cells after evaluation.

- the anchor formula cell keeps the formula text
- the anchor cell stores the spill reference such as `A1:B3`
- spilled cells receive cached values
- stale spill cells are cleared if the result shrinks or the query becomes invalid

That persisted spill output is what survives `SaveAs(...)` and `OpenFile(...)`.

## Data Mapping Rules

When a worksheet is copied into PostgreSQL:

- row 1 becomes the SQL column header row
- blank headers are replaced with generated names such as `_col_A`
- duplicate headers are made unique, for example `Amount__2`
- missing cells become `NULL`
- present worksheet values are inserted as text
- compatibility helpers provide workbook-style `sum(text)`, `avg(text)`, `instr(...)`, and permissive `CAST(... AS REAL|INTEGER)` behavior

This gives SQL a table-like view of a worksheet without changing the original sheet data model.

## Source Resolution

By default, source tokens in `FROM` or `JOIN` are treated as worksheet names.

Example:

```sql
select * from "Data"
```

The engine also supports custom source resolution through `SetSQLSourceResolver(...)`. This is used for non-standard identifiers such as `gid_7`.

Example mapping:

- `gid_7` -> `Sales`
- `gid_0` -> `Traffic`

That allows upstream systems to expose stable logical identifiers without forcing worksheet names to appear directly in formulas.

## Dependency Behavior

SQL formulas integrate with the workbook dependency system, but the dependency granularity is coarse.

The engine marks a SQL formula as dependent on whole source sheets, not exact referenced cells.

That means:

- any relevant change in a source sheet can trigger recalculation
- dependency tracking is simpler and safer
- recalculation scope may be larger than strictly necessary

This is a deliberate tradeoff because parsing arbitrary SQL into exact cell-level dependencies would be much more complex and error-prone.

## Why Use SQL Instead of Normal Excel Formulas

SQL is a good fit when the output is tabular and relational:

- summarizing rows by category
- filtering messy imported data
- building grouped report tables
- doing joins or derived-table logic
- cleaning text before aggregation

Normal Excel formulas are still a better fit when:

- the logic is cell-oriented rather than table-oriented
- the result is scalar and local
- users need Excel-native semantics and compatibility first
- the operation depends heavily on spreadsheet-specific functions

## Pros

- expressive for analytical and reporting workloads
- much easier to read than deeply nested lookup and aggregation formulas
- supports `WITH`, subqueries, `CASE`, `CAST`, and standard SQL grouping/filtering
- workbook-aware source resolution keeps the query inside the Excel engine
- returns a spill matrix naturally from one anchor formula
- persisted spill results survive save and reopen
- read-only restriction reduces risk from destructive queries

## Cons

- every evaluation rebuilds temporary PostgreSQL tables from worksheet rows
- dependency tracking is sheet-level, not cell-level
- workbook-style type compatibility is intentionally limited to the helper functions installed for SQL formulas
- users need to know SQL syntax in addition to spreadsheet formulas
- exact quoted identifier matching can be strict for messy headers
- large sheets can make repeated SQL recalculation expensive

## Why PostgreSQL Is Used

PostgreSQL is used as the SQL runtime for workbook-aware formulas.

It is a practical choice because it provides:

- support for `SELECT`, `WITH`, grouping, sorting, casting, and expressions
- process isolation from this Go library
- a shared execution model with the upstream `excelize-mcp` PostgreSQL sheet engine
- a DSN-based deployment model for standalone database processes

The design goal is not persistent storage. The goal is to get a compact embedded SQL runtime over workbook data.

## Error and Spill Management

SQL formulas have extra persistence logic because the result can span many cells.

The engine must:

- calculate the matrix result
- update the anchor formula spill reference
- write all spilled values
- clear old spill cells if the new result is smaller
- clear old spill cells if the query becomes invalid
- refresh worksheet dimensions after spill changes

This is why SQL formulas use a spill-aware persistence path instead of the normal scalar formula cache path.

### Spill Lifecycle

```mermaid
stateDiagram-v2
    [*] --> Evaluating
    Evaluating --> Spilled: query succeeds
    Evaluating --> ErrorState: query fails
    Spilled --> Resized: result shape changes
    Resized --> Spilled
    Spilled --> ErrorState: formula changed to invalid SQL
    ErrorState --> Evaluating: formula recalculated
```

## Practical Summary

`SQL("...")` is best understood as:

- a native formula entry point
- backed by temporary PostgreSQL tables over a DSN connection
- sourcing data from workbook worksheets
- producing a persistent spill range in the workbook

That combination gives the project a relational query capability while removing the embedded SQLite dependency from this repo.

## Relevant Implementation Files

- `sql_formula.go`
- `calc.go`
- `batch_dag_scheduler.go`
- `excelize.go`
- `sql_formula_test.go`
- `calc_sql_formula_persist_test.go`
