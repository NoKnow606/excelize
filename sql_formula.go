package excelize

import (
	"container/list"
	"database/sql"
	"errors"
	"fmt"
	"regexp"
	"strconv"
	"strings"
	"unicode"

	_ "modernc.org/sqlite"
)

var (
	errSQLFormulaEmptyQuery          = errors.New("SQL formula query cannot be empty")
	errSQLFormulaOnlySelectSupported = errors.New("only single SELECT statements are supported")
)

// SQLSourceResolver resolves SQL source tokens such as worksheet names or gid_*
// into workbook sheet names.
type SQLSourceResolver interface {
	ResolveSQLSource(token string, sheetList []string) (string, error)
}

// SQLSourceResolverFunc adapts a function into SQLSourceResolver.
type SQLSourceResolverFunc func(token string, sheetList []string) (string, error)

// ResolveSQLSource implements SQLSourceResolver.
func (fn SQLSourceResolverFunc) ResolveSQLSource(token string, sheetList []string) (string, error) {
	return fn(token, sheetList)
}

// SQLQueryResult contains the SQL result matrix, including the header row.
type SQLQueryResult struct {
	Columns     []string
	Matrix      [][]interface{}
	SourceSheet string
}

// SQLSourceBinding reports which worksheet a SQL source token resolved to.
type SQLSourceBinding struct {
	WorksheetName string
	TableName     string
}

// SQLCompileResult contains the normalized query and rewritten table bindings.
type SQLCompileResult struct {
	Query        string
	RewrittenSQL string
	SourceSheet  string
	Sources      []SQLSourceBinding
}

type sqlQuerySource struct {
	Start int
	End   int
	Token string
}

type sqlResolvedSource struct {
	SheetName string
	TableName string
}

type preparedSQL struct {
	query          string
	rewrittenQuery string
	sources        []sqlResolvedSource
	sourceSheet    string
	db             *sql.DB
}

// IsSQLFormula reports whether the input is a SQL formula.
func IsSQLFormula(formula string) bool {
	trimmed := strings.TrimSpace(formula)
	trimmed = strings.TrimPrefix(trimmed, "=")
	return strings.HasPrefix(strings.ToUpper(trimmed), "SQL(")
}

// CompileSQL compiles a SQL formula or raw SQL query against workbook sheets.
func (f *File) CompileSQL(sqlInput string) (*SQLCompileResult, error) {
	prepared, err := f.prepareSQL(sqlInput)
	if err != nil {
		return nil, err
	}
	defer prepared.db.Close()

	return &SQLCompileResult{
		Query:        prepared.query,
		RewrittenSQL: prepared.rewrittenQuery,
		SourceSheet:  prepared.sourceSheet,
		Sources:      buildSQLSourceBindings(prepared.sources),
	}, nil
}

// ExecuteSQL executes a SQL formula or raw SQL query against workbook sheets.
func (f *File) ExecuteSQL(sqlInput string) (*SQLQueryResult, error) {
	prepared, err := f.prepareSQL(sqlInput)
	if err != nil {
		return nil, err
	}
	defer prepared.db.Close()

	rows, err := prepared.db.Query(prepared.rewrittenQuery)
	if err != nil {
		return nil, fmt.Errorf("execute SQL query: %w", err)
	}
	defer rows.Close()

	columns, err := rows.Columns()
	if err != nil {
		return nil, fmt.Errorf("read query columns: %w", err)
	}

	matrix := make([][]interface{}, 0, 8)
	headerRow := make([]interface{}, len(columns))
	for i, column := range columns {
		headerRow[i] = column
	}
	matrix = append(matrix, headerRow)

	for rows.Next() {
		values := make([]interface{}, len(columns))
		scanTargets := make([]interface{}, len(columns))
		for i := range values {
			scanTargets[i] = &values[i]
		}
		if err := rows.Scan(scanTargets...); err != nil {
			return nil, fmt.Errorf("scan query row: %w", err)
		}

		row := make([]interface{}, len(columns))
		for i, value := range values {
			row[i] = normalizeSQLiteValue(value)
		}
		matrix = append(matrix, row)
	}

	if err := rows.Err(); err != nil {
		return nil, fmt.Errorf("iterate SQL query rows: %w", err)
	}

	return &SQLQueryResult{
		Columns:     columns,
		Matrix:      matrix,
		SourceSheet: prepared.sourceSheet,
	}, nil
}

// ExecuteSQLFormula is a convenience wrapper for formula-style SQL input.
func (f *File) ExecuteSQLFormula(formula string) (*SQLQueryResult, error) {
	return f.ExecuteSQL(formula)
}

// SQL executes workbook-aware SQL and returns the result as a spill matrix.
func (fn *formulaFuncs) SQL(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "SQL requires exactly 1 argument")
	}
	queryArg := argsList.Front().Value.(formulaArg)
	if queryArg.Type == ArgError {
		return queryArg
	}

	result, err := fn.f.ExecuteSQL(queryArg.Value())
	if err != nil {
		return newErrorFormulaArg(formulaErrorVALUE, err.Error())
	}
	return sqlQueryResultToFormulaArg(result)
}

func sqlQueryResultToFormulaArg(result *SQLQueryResult) formulaArg {
	if result == nil || len(result.Matrix) == 0 {
		return newMatrixFormulaArg([][]formulaArg{{newEmptyFormulaArg()}})
	}
	matrix := make([][]formulaArg, len(result.Matrix))
	for r := range result.Matrix {
		matrix[r] = make([]formulaArg, len(result.Matrix[r]))
		for c := range result.Matrix[r] {
			matrix[r][c] = interfaceToFormulaArg(result.Matrix[r][c])
		}
	}
	return newMatrixFormulaArg(matrix)
}

func interfaceToFormulaArg(value interface{}) formulaArg {
	switch typed := value.(type) {
	case nil:
		return newEmptyFormulaArg()
	case formulaArg:
		return typed
	case string:
		return newStringFormulaArg(typed)
	case []byte:
		return newStringFormulaArg(string(typed))
	case bool:
		return newBoolFormulaArg(typed)
	case int:
		return newNumberFormulaArg(float64(typed))
	case int8:
		return newNumberFormulaArg(float64(typed))
	case int16:
		return newNumberFormulaArg(float64(typed))
	case int32:
		return newNumberFormulaArg(float64(typed))
	case int64:
		return newNumberFormulaArg(float64(typed))
	case uint:
		return newNumberFormulaArg(float64(typed))
	case uint8:
		return newNumberFormulaArg(float64(typed))
	case uint16:
		return newNumberFormulaArg(float64(typed))
	case uint32:
		return newNumberFormulaArg(float64(typed))
	case uint64:
		return newNumberFormulaArg(float64(typed))
	case float32:
		return newNumberFormulaArg(float64(typed))
	case float64:
		return newNumberFormulaArg(typed)
	default:
		return newStringFormulaArg(fmt.Sprint(value))
	}
}

func (f *File) prepareSQL(sqlInput string) (*preparedSQL, error) {
	query, err := extractSQLInput(sqlInput)
	if err != nil {
		return nil, err
	}

	rewrittenQuery, sources, cteNames, err := rewriteQuerySources(query, f.GetSheetList(), f.sqlSourceResolver)
	if err != nil {
		return nil, err
	}
	sourceSheet := ""
	if len(sources) > 0 {
		sourceSheet = sources[0].SheetName
	}

	db, err := sql.Open("sqlite", ":memory:")
	if err != nil {
		return nil, fmt.Errorf("open in-memory sqlite: %w", err)
	}

	headersByTable, err := materializeSheets(db, f, sources)
	if err != nil {
		db.Close()
		return nil, err
	}
	if err := validateQuotedIdentifiers(rewrittenQuery, headersByTable, cteNames); err != nil {
		db.Close()
		return nil, err
	}
	stmt, err := db.Prepare(rewrittenQuery)
	if err != nil {
		db.Close()
		return nil, fmt.Errorf("compile SQL query: %w", err)
	}
	stmt.Close()

	return &preparedSQL{
		query:          query,
		rewrittenQuery: rewrittenQuery,
		sources:        sources,
		sourceSheet:    sourceSheet,
		db:             db,
	}, nil
}

func (f *File) sqlFormulaSourceSheets(formula string) []string {
	sheets, err := extractSQLSourceSheets(formula, f.GetSheetList(), f.sqlSourceResolver)
	if err != nil {
		return nil
	}
	return sheets
}

func (f *File) extractDependenciesWithSQL(formula, currentSheet, currentCell string) []string {
	if isExternalCachedFormula(formula) {
		return nil
	}
	deps := make(map[string]bool)
	for _, dep := range extractDependencies(formula, currentSheet, currentCell) {
		deps[dep] = true
	}
	f.appendSQLDependencyMarkers(deps, formula, nil)
	return dependencyMapToSlice(deps)
}

func (f *File) extractDependenciesOptimizedWithSQL(formula, currentSheet, currentCell string, columnIndex map[string][]string, columnMetadata map[string]*columnMeta) []string {
	if isExternalCachedFormula(formula) {
		return nil
	}
	deps := make(map[string]bool)
	for _, dep := range extractDependenciesOptimized(formula, currentSheet, currentCell, columnIndex, columnMetadata) {
		deps[dep] = true
	}
	f.appendSQLDependencyMarkers(deps, formula, columnMetadata)
	return dependencyMapToSlice(deps)
}

func (f *File) appendSQLDependencyMarkers(deps map[string]bool, formula string, columnMetadata map[string]*columnMeta) {
	for _, sheetName := range f.sqlFormulaSourceSheets(formula) {
		deps["SHEET:"+sheetName] = true
		if columnMetadata == nil {
			continue
		}
		for colKey, meta := range columnMetadata {
			if meta.hasFormulas && strings.HasPrefix(colKey, sheetName+"!") {
				deps["COLUMN:"+colKey] = true
			}
		}
	}
}

func dependencyMapToSlice(deps map[string]bool) []string {
	result := make([]string, 0, len(deps))
	for dep := range deps {
		result = append(result, dep)
	}
	return result
}

func extractSQLSourceSheets(sqlInput string, sheetList []string, resolver SQLSourceResolver) ([]string, error) {
	query, err := extractSQLInput(sqlInput)
	if err != nil {
		return nil, err
	}

	cteNames, err := collectCTENames(query)
	if err != nil {
		return nil, err
	}
	sourceTokens, err := findSourceTokens(query)
	if err != nil {
		return nil, err
	}

	seen := make(map[string]bool, len(sourceTokens))
	sheets := make([]string, 0, len(sourceTokens))
	for _, sourceToken := range sourceTokens {
		identifier := strings.ToLower(unquoteIdentifier(sourceToken.Token))
		if _, isCTE := cteNames[identifier]; isCTE {
			continue
		}
		sheetName, err := resolveSQLSourceName(sourceToken.Token, sheetList, resolver)
		if err != nil {
			return nil, err
		}
		if !seen[sheetName] {
			seen[sheetName] = true
			sheets = append(sheets, sheetName)
		}
	}
	return sheets, nil
}

func extractQuery(formula string) (string, error) {
	trimmed := strings.TrimSpace(formula)
	trimmed = strings.TrimPrefix(trimmed, "=")

	openIdx := strings.Index(trimmed, "(")
	closeIdx := strings.LastIndex(trimmed, ")")
	if openIdx < 0 || closeIdx <= openIdx {
		return "", fmt.Errorf("invalid SQL formula syntax")
	}

	fnName := strings.TrimSpace(trimmed[:openIdx])
	if !strings.EqualFold(fnName, "SQL") {
		return "", fmt.Errorf("unsupported formula function %q", fnName)
	}

	inner := strings.TrimSpace(trimmed[openIdx+1 : closeIdx])
	if inner == "" {
		return "", errSQLFormulaEmptyQuery
	}

	query, remainder, err := parseExcelStringLiteral(inner)
	if err != nil {
		return "", err
	}
	if strings.TrimSpace(remainder) != "" {
		return "", fmt.Errorf("SQL formula currently accepts exactly one string argument")
	}

	query = validateSQLQueryText(query)
	if query == "" {
		return "", errSQLFormulaEmptyQuery
	}
	if strings.Contains(query, ";") {
		return "", errSQLFormulaOnlySelectSupported
	}
	if !startsWithSelectOrWith(query) {
		return "", errSQLFormulaOnlySelectSupported
	}

	return query, nil
}

func extractSQLInput(sqlInput string) (string, error) {
	trimmed := strings.TrimSpace(sqlInput)
	trimmed = strings.TrimPrefix(trimmed, "=")
	if strings.HasPrefix(strings.ToUpper(trimmed), "SQL(") {
		return extractQuery(sqlInput)
	}

	query := validateSQLQueryText(sqlInput)
	if query == "" {
		return "", errSQLFormulaEmptyQuery
	}
	if strings.Contains(query, ";") {
		return "", errSQLFormulaOnlySelectSupported
	}
	if !startsWithSelectOrWith(query) {
		return "", errSQLFormulaOnlySelectSupported
	}

	return query, nil
}

func validateSQLQueryText(query string) string {
	query = strings.TrimSpace(query)
	query = strings.TrimSpace(strings.TrimSuffix(query, ";"))
	return query
}

func parseExcelStringLiteral(input string) (string, string, error) {
	if input == "" || input[0] != '"' {
		return "", "", fmt.Errorf("SQL formula query must be wrapped in double quotes")
	}

	var builder strings.Builder
	for i := 1; i < len(input); i++ {
		ch := input[i]
		if ch == '"' {
			if i+1 < len(input) && input[i+1] == '"' {
				builder.WriteByte('"')
				i++
				continue
			}
			return builder.String(), input[i+1:], nil
		}
		builder.WriteByte(ch)
	}

	return "", "", fmt.Errorf("unterminated SQL string literal in formula")
}

func rewriteQuerySources(
	query string,
	sheetList []string,
	resolver SQLSourceResolver,
) (string, []sqlResolvedSource, map[string]struct{}, error) {
	cteNames, err := collectCTENames(query)
	if err != nil {
		return "", nil, nil, err
	}

	sourceTokens, err := findSourceTokens(query)
	if err != nil {
		return "", nil, nil, err
	}

	sourcesBySheet := make(map[string]sqlResolvedSource, len(sourceTokens))
	orderedSources := make([]sqlResolvedSource, 0, len(sourceTokens))
	replacements := make([]*sqlResolvedSource, 0, len(sourceTokens))

	for _, sourceToken := range sourceTokens {
		identifier := strings.ToLower(unquoteIdentifier(sourceToken.Token))
		if _, isCTE := cteNames[identifier]; isCTE {
			replacements = append(replacements, nil)
			continue
		}

		sourceName, err := resolveSQLSourceName(sourceToken.Token, sheetList, resolver)
		if err != nil {
			return "", nil, nil, err
		}

		source, ok := sourcesBySheet[sourceName]
		if !ok {
			source = sqlResolvedSource{
				SheetName: sourceName,
				TableName: fmt.Sprintf("__sheet_source_%d__", len(orderedSources)),
			}
			sourcesBySheet[sourceName] = source
			orderedSources = append(orderedSources, source)
		}
		replacement := sqlResolvedSource{
			SheetName: sourceToken.Token,
			TableName: source.TableName,
		}
		replacements = append(replacements, &replacement)
	}

	var builder strings.Builder
	last := 0
	for idx, sourceToken := range sourceTokens {
		builder.WriteString(query[last:sourceToken.Start])
		if replacements[idx] == nil {
			builder.WriteString(query[sourceToken.Start:sourceToken.End])
		} else {
			builder.WriteString(`"`)
			builder.WriteString(replacements[idx].TableName)
			builder.WriteString(`"`)
		}
		last = sourceToken.End
	}
	builder.WriteString(query[last:])

	return builder.String(), orderedSources, cteNames, nil
}

func findSourceTokens(query string) ([]sqlQuerySource, error) {
	tokens, foundFrom, err := findSourceTokensInRange(query, 0, len(query))
	if err != nil {
		return nil, err
	}
	if !foundFrom {
		return nil, fmt.Errorf("missing FROM clause")
	}
	return tokens, nil
}

func findSourceTokensInRange(query string, start, end int) ([]sqlQuerySource, bool, error) {
	inSingle := false
	inDouble := false
	inBacktick := false
	foundFrom := false
	tokens := make([]sqlQuerySource, 0, 4)

	for i := start; i < end; i++ {
		ch := query[i]

		switch ch {
		case '\'':
			if !inDouble && !inBacktick {
				inSingle = !inSingle
			}
		case '"':
			if !inSingle && !inBacktick {
				inDouble = !inDouble
			}
		case '`':
			if !inSingle && !inDouble {
				inBacktick = !inBacktick
			}
		}

		if inSingle || inDouble || inBacktick {
			continue
		}
		if ch == '(' {
			matchIdx, err := findMatchingParenInRange(query, i, end)
			if err != nil {
				return nil, false, err
			}
			nestedTokens, nestedFoundFrom, err := findSourceTokensInRange(query, i+1, matchIdx)
			if err != nil {
				return nil, false, err
			}
			tokens = append(tokens, nestedTokens...)
			foundFrom = foundFrom || nestedFoundFrom
			i = matchIdx
			continue
		}

		keyword := ""
		switch {
		case hasKeywordAt(query, i, "from"):
			keyword = "from"
			foundFrom = true
		case hasKeywordAt(query, i, "join"):
			keyword = "join"
		default:
			continue
		}

		j := i + len(keyword)
		for j < end && unicode.IsSpace(rune(query[j])) {
			j++
		}
		if j >= end {
			return nil, false, fmt.Errorf("missing source after %s", strings.ToUpper(keyword))
		}
		if query[j] == '(' {
			matchIdx, err := findMatchingParenInRange(query, j, end)
			if err != nil {
				return nil, false, err
			}
			nestedTokens, nestedFoundFrom, err := findSourceTokensInRange(query, j+1, matchIdx)
			if err != nil {
				return nil, false, err
			}
			tokens = append(tokens, nestedTokens...)
			foundFrom = foundFrom || nestedFoundFrom
			i = matchIdx
			continue
		}

		start := j
		if isIdentifierQuote(query[j]) {
			quote := query[j]
			j++
			for j < end {
				if query[j] == quote {
					if j+1 < len(query) && query[j+1] == quote {
						j += 2
						continue
					}
					tokens = append(tokens, sqlQuerySource{
						Start: start,
						End:   j + 1,
						Token: query[start : j+1],
					})
					i = j
					goto nextToken
				}
				j++
			}
			return nil, false, fmt.Errorf("unterminated quoted source after %s", strings.ToUpper(keyword))
		}

		for j < end {
			if unicode.IsSpace(rune(query[j])) || query[j] == ',' || query[j] == ')' {
				break
			}
			j++
		}
		tokens = append(tokens, sqlQuerySource{
			Start: start,
			End:   j,
			Token: query[start:j],
		})
		i = j - 1

	nextToken:
	}

	if !foundFrom {
		return tokens, false, nil
	}
	return tokens, true, nil
}

func resolveSQLSourceName(token string, sheetList []string, resolver SQLSourceResolver) (string, error) {
	identifier := unquoteIdentifier(token)
	if identifier == "" {
		return "", fmt.Errorf("empty source in FROM clause")
	}

	lowerIdentifier := strings.ToLower(identifier)
	if resolver != nil {
		if resolved, err := resolver.ResolveSQLSource(token, sheetList); err == nil {
			return resolved, nil
		} else if strings.HasPrefix(lowerIdentifier, "gid_") {
			return "", err
		}
	}

	if strings.HasPrefix(lowerIdentifier, "gid_") {
		return "", fmt.Errorf("worksheet for %s was not found", identifier)
	}

	for _, sheetName := range sheetList {
		if sheetName == identifier {
			return sheetName, nil
		}
	}

	return "", fmt.Errorf("worksheet %q was not found", identifier)
}

func materializeSheets(db *sql.DB, f *File, sources []sqlResolvedSource) (map[string][]string, error) {
	headersByTable := make(map[string][]string, len(sources))
	for _, source := range sources {
		if _, ok := headersByTable[source.TableName]; ok {
			continue
		}

		headers, err := materializeSheet(db, f, source.TableName, source.SheetName)
		if err != nil {
			return nil, err
		}
		headersByTable[source.TableName] = headers
	}

	return headersByTable, nil
}

func buildSQLSourceBindings(sources []sqlResolvedSource) []SQLSourceBinding {
	bindings := make([]SQLSourceBinding, 0, len(sources))
	for _, source := range sources {
		bindings = append(bindings, SQLSourceBinding{
			WorksheetName: source.SheetName,
			TableName:     source.TableName,
		})
	}
	return bindings
}

func materializeSheet(db *sql.DB, f *File, tableName string, sheetName string) ([]string, error) {
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("read source worksheet %q: %w", sheetName, err)
	}
	if len(rows) == 0 {
		return nil, fmt.Errorf("source worksheet %q is empty", sheetName)
	}

	headers := buildSQLHeaders(rows)
	if len(headers) == 0 {
		return nil, fmt.Errorf("source worksheet %q has no usable columns", sheetName)
	}

	createSQL := buildCreateTableSQL(tableName, headers)
	if _, err := db.Exec(createSQL); err != nil {
		return nil, fmt.Errorf("create in-memory SQL table: %w", err)
	}

	placeholders := strings.TrimSuffix(strings.Repeat("?,", len(headers)), ",")
	insertSQL := fmt.Sprintf(
		`INSERT INTO "%s" (%s) VALUES (%s)`,
		tableName,
		buildQuotedIdentifierList(headers),
		placeholders,
	)

	tx, err := db.Begin()
	if err != nil {
		return nil, fmt.Errorf("begin SQL transaction: %w", err)
	}
	defer tx.Rollback()

	stmt, err := tx.Prepare(insertSQL)
	if err != nil {
		return nil, fmt.Errorf("prepare SQL insert: %w", err)
	}
	defer stmt.Close()

	for _, row := range rows[1:] {
		values := make([]interface{}, len(headers))
		for colIdx := range headers {
			switch {
			case colIdx >= len(row):
				values[colIdx] = nil
			case row[colIdx] == "":
				values[colIdx] = ""
			default:
				values[colIdx] = coerceCellValue(row[colIdx])
			}
		}
		if _, err := stmt.Exec(values...); err != nil {
			return nil, fmt.Errorf("insert worksheet row into SQL table: %w", err)
		}
	}

	if err := tx.Commit(); err != nil {
		return nil, fmt.Errorf("commit SQL transaction: %w", err)
	}
	return headers, nil
}

func buildSQLHeaders(rows [][]string) []string {
	maxCols := 0
	for _, row := range rows {
		if len(row) > maxCols {
			maxCols = len(row)
		}
	}

	seen := make(map[string]int)
	headers := make([]string, 0, maxCols)
	for colIdx := 0; colIdx < maxCols; colIdx++ {
		header := ""
		if colIdx < len(rows[0]) {
			header = strings.TrimSpace(rows[0][colIdx])
		}
		if header == "" {
			header = fmt.Sprintf("_col_%s", sqlColumnIndexToLetter(colIdx))
		}

		key := strings.ToLower(header)
		if count := seen[key]; count > 0 {
			header = fmt.Sprintf("%s__%d", header, count+1)
			key = strings.ToLower(header)
		}
		seen[key]++
		headers = append(headers, header)
	}
	return headers
}

func validateQuotedIdentifiers(query string, headersByTable map[string][]string, cteNames map[string]struct{}) error {
	known := make(map[string]struct{}, len(headersByTable)*4)
	for tableName, headers := range headersByTable {
		known[tableName] = struct{}{}
		for _, header := range headers {
			known[header] = struct{}{}
		}
	}
	for cteName := range cteNames {
		known[cteName] = struct{}{}
	}

	aliases := collectQuotedAliases(query)
	for alias := range aliases {
		known[alias] = struct{}{}
	}

	quotedIdentifiers, err := collectDoubleQuotedTokens(query)
	if err != nil {
		return err
	}

	for _, ident := range quotedIdentifiers {
		if _, ok := known[ident]; ok {
			continue
		}
		if suggestion := findWhitespaceNormalizedMatch(ident, known); suggestion != "" {
			return fmt.Errorf(
				"quoted identifier %q does not match any column exactly. Did you mean %q?",
				ident,
				suggestion,
			)
		}

		if !strings.ContainsAny(ident, "\r\n\t") {
			continue
		}

		return fmt.Errorf(
			"quoted identifier %q does not match any column exactly; remove embedded line breaks or indentation",
			ident,
		)
	}

	return nil
}

func collectQuotedAliases(query string) map[string]struct{} {
	aliases := make(map[string]struct{})
	inSingle := false
	inDouble := false
	inBacktick := false

	for i := 0; i < len(query); i++ {
		ch := query[i]

		switch ch {
		case '\'':
			if !inDouble && !inBacktick {
				inSingle = !inSingle
			}
		case '"':
			if !inSingle && !inBacktick {
				inDouble = !inDouble
			}
		case '`':
			if !inSingle && !inDouble {
				inBacktick = !inBacktick
			}
		}

		if inSingle || inDouble || inBacktick {
			continue
		}

		if !hasKeywordAt(query, i, "as") {
			continue
		}

		j := i + len("as")
		for j < len(query) && unicode.IsSpace(rune(query[j])) {
			j++
		}
		if j >= len(query) || query[j] != '"' {
			continue
		}

		start := j
		j++
		for j < len(query) {
			if query[j] == '"' {
				if j+1 < len(query) && query[j+1] == '"' {
					j += 2
					continue
				}
				aliases[unquoteIdentifier(query[start:j+1])] = struct{}{}
				i = j
				break
			}
			j++
		}
	}

	return aliases
}

func collectDoubleQuotedTokens(query string) ([]string, error) {
	inSingle := false
	inBacktick := false
	tokens := make([]string, 0, 8)

	for i := 0; i < len(query); i++ {
		ch := query[i]
		switch ch {
		case '\'':
			if !inBacktick {
				inSingle = !inSingle
			}
		case '`':
			if !inSingle {
				inBacktick = !inBacktick
			}
		case '"':
			if inSingle || inBacktick {
				continue
			}
			start := i
			i++
			for i < len(query) {
				if query[i] == '"' {
					if i+1 < len(query) && query[i+1] == '"' {
						i += 2
						continue
					}
					tokens = append(tokens, unquoteIdentifier(query[start:i+1]))
					break
				}
				i++
			}
			if i >= len(query) {
				return nil, fmt.Errorf("unterminated quoted identifier in SQL query")
			}
		}
	}

	return tokens, nil
}

func findWhitespaceNormalizedMatch(identifier string, known map[string]struct{}) string {
	if normalized := removeWhitespace(identifier); normalized != identifier {
		for candidate := range known {
			if removeWhitespace(candidate) == normalized {
				return candidate
			}
		}
	}

	if collapsed := collapseWhitespace(identifier); collapsed != identifier {
		for candidate := range known {
			if collapseWhitespace(candidate) == collapsed {
				return candidate
			}
		}
	}

	return ""
}

func removeWhitespace(input string) string {
	var builder strings.Builder
	for _, r := range input {
		if unicode.IsSpace(r) {
			continue
		}
		builder.WriteRune(r)
	}
	return builder.String()
}

func collapseWhitespace(input string) string {
	var builder strings.Builder
	lastWasSpace := false
	for _, r := range input {
		if unicode.IsSpace(r) {
			if builder.Len() == 0 || lastWasSpace {
				lastWasSpace = true
				continue
			}
			builder.WriteByte(' ')
			lastWasSpace = true
			continue
		}
		builder.WriteRune(r)
		lastWasSpace = false
	}
	return strings.TrimSpace(builder.String())
}

func buildCreateTableSQL(tableName string, headers []string) string {
	columns := make([]string, len(headers))
	for i, header := range headers {
		columns[i] = fmt.Sprintf(`"%s"`, escapeDoubleQuotes(header))
	}
	return fmt.Sprintf(`CREATE TABLE "%s" (%s)`, tableName, strings.Join(columns, ", "))
}

func buildQuotedIdentifierList(headers []string) string {
	quoted := make([]string, len(headers))
	for i, header := range headers {
		quoted[i] = fmt.Sprintf(`"%s"`, escapeDoubleQuotes(header))
	}
	return strings.Join(quoted, ", ")
}

func coerceCellValue(raw string) interface{} {
	trimmed := strings.TrimSpace(raw)
	switch strings.ToLower(trimmed) {
	case "true":
		return true
	case "false":
		return false
	}

	if isLikelyInteger(trimmed) {
		if intVal, err := strconv.ParseInt(trimmed, 10, 64); err == nil {
			return intVal
		}
	}
	if floatVal, err := strconv.ParseFloat(trimmed, 64); err == nil {
		return floatVal
	}

	return raw
}

func isLikelyInteger(value string) bool {
	if value == "" {
		return false
	}
	if matched, _ := regexp.MatchString(`^-?\d+$`, value); !matched {
		return false
	}

	unsigned := strings.TrimPrefix(value, "-")
	if len(unsigned) > 1 && unsigned[0] == '0' {
		return false
	}
	return len(unsigned) <= 15
}

func normalizeSQLiteValue(value interface{}) interface{} {
	switch typed := value.(type) {
	case nil:
		return nil
	case []byte:
		return string(typed)
	default:
		return typed
	}
}

func startsWithSelectOrWith(query string) bool {
	trimmed := strings.TrimSpace(query)
	return strings.HasPrefix(strings.ToUpper(trimmed), "SELECT ") ||
		strings.HasPrefix(strings.ToUpper(trimmed), "WITH ")
}

func hasKeywordAt(input string, index int, keyword string) bool {
	if index < 0 || index+len(keyword) > len(input) {
		return false
	}
	if !strings.EqualFold(input[index:index+len(keyword)], keyword) {
		return false
	}

	beforeOK := index == 0 || !isWordChar(rune(input[index-1]))
	afterIdx := index + len(keyword)
	afterOK := afterIdx >= len(input) || !isWordChar(rune(input[afterIdx]))
	return beforeOK && afterOK
}

func collectCTENames(query string) (map[string]struct{}, error) {
	cteNames := make(map[string]struct{})
	trimmed := strings.TrimSpace(query)
	if !strings.HasPrefix(strings.ToUpper(trimmed), "WITH ") {
		return cteNames, nil
	}

	pos := len("WITH")
	pos = skipSQLWhitespace(trimmed, pos)
	if hasKeywordAt(trimmed, pos, "RECURSIVE") {
		pos += len("RECURSIVE")
	}

	for pos < len(trimmed) {
		pos = skipSQLWhitespace(trimmed, pos)
		start := pos
		if start >= len(trimmed) {
			return nil, fmt.Errorf("invalid WITH clause")
		}

		if isIdentifierQuote(trimmed[start]) {
			quote := trimmed[start]
			pos++
			for pos < len(trimmed) {
				if trimmed[pos] == quote {
					if pos+1 < len(trimmed) && trimmed[pos+1] == quote {
						pos += 2
						continue
					}
					cteNames[strings.ToLower(unquoteIdentifier(trimmed[start:pos+1]))] = struct{}{}
					pos++
					break
				}
				pos++
			}
		} else {
			for pos < len(trimmed) && isWordChar(rune(trimmed[pos])) {
				pos++
			}
			if pos == start {
				return nil, fmt.Errorf("invalid CTE name in WITH clause")
			}
			cteNames[strings.ToLower(trimmed[start:pos])] = struct{}{}
		}

		pos = skipSQLWhitespace(trimmed, pos)
		if !hasKeywordAt(trimmed, pos, "AS") {
			return nil, fmt.Errorf("expected AS in WITH clause")
		}
		pos += len("AS")
		pos = skipSQLWhitespace(trimmed, pos)
		if pos >= len(trimmed) || trimmed[pos] != '(' {
			return nil, fmt.Errorf("expected opening parenthesis after AS")
		}
		endPos, err := findMatchingParen(trimmed, pos)
		if err != nil {
			return nil, err
		}
		pos = endPos + 1
		pos = skipSQLWhitespace(trimmed, pos)
		if pos >= len(trimmed) || trimmed[pos] != ',' {
			break
		}
		pos++
	}

	return cteNames, nil
}

func skipSQLWhitespace(input string, pos int) int {
	for pos < len(input) && unicode.IsSpace(rune(input[pos])) {
		pos++
	}
	return pos
}

func findMatchingParen(input string, start int) (int, error) {
	return findMatchingParenInRange(input, start, len(input))
}

func findMatchingParenInRange(input string, start, end int) (int, error) {
	depth := 0
	inSingle := false
	inDouble := false
	inBacktick := false

	for i := start; i < end; i++ {
		ch := input[i]
		switch ch {
		case '\'':
			if !inDouble && !inBacktick {
				inSingle = !inSingle
			}
		case '"':
			if !inSingle && !inBacktick {
				inDouble = !inDouble
			}
		case '`':
			if !inSingle && !inDouble {
				inBacktick = !inBacktick
			}
		}

		if inSingle || inDouble || inBacktick {
			continue
		}

		switch ch {
		case '(':
			depth++
		case ')':
			depth--
			if depth == 0 {
				return i, nil
			}
		}
	}
	return 0, fmt.Errorf("unterminated parenthesis in SQL query")
}

func isIdentifierQuote(ch byte) bool {
	return ch == '"' || ch == '`'
}

func isWordChar(r rune) bool {
	return r == '_' || unicode.IsLetter(r) || unicode.IsDigit(r)
}

func unquoteIdentifier(identifier string) string {
	identifier = strings.TrimSpace(identifier)
	if len(identifier) >= 2 && identifier[0] == '"' && identifier[len(identifier)-1] == '"' {
		return strings.ReplaceAll(identifier[1:len(identifier)-1], `""`, `"`)
	}
	if len(identifier) >= 2 && identifier[0] == '`' && identifier[len(identifier)-1] == '`' {
		return strings.ReplaceAll(identifier[1:len(identifier)-1], "``", "`")
	}
	return identifier
}

func escapeDoubleQuotes(input string) string {
	return strings.ReplaceAll(input, `"`, `""`)
}

func sqlColumnIndexToLetter(idx int) string {
	name, err := ColumnNumberToName(idx + 1)
	if err != nil {
		return fmt.Sprintf("COL%d", idx+1)
	}
	return name
}
