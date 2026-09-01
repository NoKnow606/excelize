package excelize

import (
	"fmt"
	"strings"
)

// FormulaToSQLOptions contains options for SQL generation.
type FormulaToSQLOptions struct {
	SourceTable string
}

// ConvertFormulaToSQL converts an Excel formula to a SQL SELECT statement.
func (f *File) ConvertFormulaToSQL(formula string, opts FormulaToSQLOptions) (string, error) {
	if opts.SourceTable == "" {
		opts.SourceTable = "Sheet1"
	}
	formula = strings.TrimPrefix(formula, "=")
	parser := NewFormulaParser(formula)
	ast, err := parser.Parse()
	if err != nil {
		return "", fmt.Errorf("parse error: %w", err)
	}
	gen := NewSQLGenerator(opts.SourceTable)
	return gen.Generate(ast)
}

// ============================================================
// TOKENIZER
// ============================================================

type FormulaTokenType int

const (
	TokenNumber FormulaTokenType = iota
	TokenString
	TokenCellRef
	TokenRange
	TokenFunction
	TokenOperator
	TokenLParen
	TokenRParen
	TokenComma
	TokenColon
	TokenBang
	TokenDollar
	TokenEQ
	TokenNE
	TokenLT
	TokenGT
	TokenLTE
	TokenGTE
	TokenAmpersand
	TokenAnd
	TokenOr
	TokenMinus
	TokenPlus
	TokenStar
	TokenSlash
	TokenPercent
	TokenCaret
	TokenTRUE
	TokenFALSE
	TokenEOF
)

type FormulaToken struct {
	Type  FormulaTokenType
	Value string
}

type FormulaTokenizer struct {
	input string
	pos   int
}

func NewFormulaTokenizer(input string) *FormulaTokenizer {
	return &FormulaTokenizer{input: input, pos: 0}
}

func (t *FormulaTokenizer) peek() rune {
	if t.pos >= len(t.input) {
		return 0
	}
	return rune(t.input[t.pos])
}

func (t *FormulaTokenizer) next() rune {
	if t.pos >= len(t.input) {
		return 0
	}
	r := rune(t.input[t.pos])
	t.pos++
	return r
}

func (t *FormulaTokenizer) skipWhitespace() {
	for {
		r := t.peek()
		if r == 0 || r == ' ' || r == '\t' || r == '\r' {
			if r == 0 {
				break
			}
			t.next()
		} else {
			break
		}
	}
}

func isLetter(r rune) bool {
	return (r >= 'A' && r <= 'Z') || (r >= 'a' && r <= 'z')
}

func isDigit(r rune) bool {
	return r >= '0' && r <= '9'
}

func isLetterOrDigit(r rune) bool {
	return isLetter(r) || isDigit(r)
}

func (t *FormulaTokenizer) readIdentifier() string {
	var result strings.Builder
	for {
		r := t.peek()
		if isLetterOrDigit(r) || r == '_' || r == '.' {
			result.WriteRune(t.next())
		} else {
			break
		}
	}
	return result.String()
}

func (t *FormulaTokenizer) readNumber() string {
	var result strings.Builder
	hasDecimal := false
	for {
		r := t.peek()
		if isDigit(r) {
			result.WriteRune(t.next())
		} else if r == '.' && !hasDecimal {
			hasDecimal = true
			result.WriteRune(t.next())
		} else {
			break
		}
	}
	return result.String()
}

func (t *FormulaTokenizer) readString() (string, error) {
	quote := t.next()
	var result strings.Builder
	for {
		r := t.next()
		if r == 0 {
			return "", fmt.Errorf("unterminated string")
		}
		if r == quote {
			break
		}
		if r == '\\' && t.peek() == quote {
			result.WriteRune(t.next())
		} else {
			result.WriteRune(r)
		}
	}
	return result.String(), nil
}

func (t *FormulaTokenizer) Next() (FormulaToken, error) {
	t.skipWhitespace()

	r := t.peek()
	if r == 0 {
		return FormulaToken{Type: TokenEOF, Value: ""}, nil
	}

	if r == '"' || r == '\'' {
		s, err := t.readString()
		if err != nil {
			return FormulaToken{}, err
		}
		return FormulaToken{Type: TokenString, Value: s}, nil
	}

	if isDigit(r) || (r == '.' && len(t.input) > t.pos && isDigit(rune(t.input[t.pos]))) {
		return FormulaToken{Type: TokenNumber, Value: t.readNumber()}, nil
	}

		if isLetter(r) || r == '$' {
			ident := t.readIdentifier()
			upper := strings.ToUpper(ident)

			if upper == "TRUE" {
				return FormulaToken{Type: TokenTRUE, Value: ident}, nil
			}
			if upper == "FALSE" {
				return FormulaToken{Type: TokenFALSE, Value: ident}, nil
			}

			// Check for range: A:B or $A:$B
			t.skipWhitespace()
			if t.peek() == ':' {
				t.next()
				t.skipWhitespace()
				if isLetter(t.peek()) || t.peek() == '$' {
					endIdent := t.readIdentifier()
					return FormulaToken{Type: TokenRange, Value: ident + ":" + endIdent}, nil
				}
				// Not a range, put the colon back (peek only moved past it)
				// Since we consumed ':', we need to handle it - it's a bare colon
				// Push back by not returning here; fall through to colon handling
				// Instead, re-parse as cell ref with trailing colon
				return FormulaToken{Type: TokenCellRef, Value: ident}, nil
			}

			t.skipWhitespace()
			if t.peek() == '(' {
				return FormulaToken{Type: TokenFunction, Value: upper}, nil
			}

			return FormulaToken{Type: TokenCellRef, Value: ident}, nil
		}

	cp := t.next()
	r2 := t.peek()

	switch cp {
	case '=':
		if r2 == '=' {
			t.next()
			return FormulaToken{Type: TokenEQ, Value: "=="}, nil
		}
		return FormulaToken{Type: TokenOperator, Value: "="}, nil
	case '<':
		if r2 == '=' {
			t.next()
			return FormulaToken{Type: TokenLTE, Value: "<="}, nil
		}
		if r2 == '>' {
			t.next()
			return FormulaToken{Type: TokenNE, Value: "<>"}, nil
		}
		return FormulaToken{Type: TokenLT, Value: "<"}, nil
	case '>':
		if r2 == '=' {
			t.next()
			return FormulaToken{Type: TokenGTE, Value: ">="}, nil
		}
		return FormulaToken{Type: TokenGT, Value: ">"}, nil
	case '&':
		return FormulaToken{Type: TokenAmpersand, Value: "&"}, nil
	case '*':
		return FormulaToken{Type: TokenStar, Value: "*"}, nil
	case '/':
		return FormulaToken{Type: TokenSlash, Value: "/"}, nil
	case '+':
		return FormulaToken{Type: TokenPlus, Value: "+"}, nil
	case '-':
		return FormulaToken{Type: TokenMinus, Value: "-"}, nil
	case '^':
		return FormulaToken{Type: TokenCaret, Value: "^"}, nil
	case '%':
		return FormulaToken{Type: TokenPercent, Value: "%"}, nil
	case '(':
		return FormulaToken{Type: TokenLParen, Value: "("}, nil
	case ')':
		return FormulaToken{Type: TokenRParen, Value: ")"}, nil
	case ',':
		return FormulaToken{Type: TokenComma, Value: ","}, nil
	case ':':
		return FormulaToken{Type: TokenColon, Value: ":"}, nil
	case '!':
		return FormulaToken{Type: TokenBang, Value: "!"}, nil
	case '$':
		return FormulaToken{Type: TokenDollar, Value: "$"}, nil
	}

	return FormulaToken{}, fmt.Errorf("unknown character: %c", cp)
}

// ============================================================
// PARSER
// ============================================================

type FormulaAST struct {
	Root Expr
}

type Expr interface {
	exprNode()
	String() string
}

func (NumberLiteral) exprNode()  {}
func (StringLiteral) exprNode()  {}
func (BooleanLiteral) exprNode() {}
func (CellRefExpr) exprNode()    {}
func (RangeExpr) exprNode()     {}
func (BinaryOpExpr) exprNode()   {}
func (UnaryOpExpr) exprNode()   {}
func (FunctionCallExpr) exprNode() {}
func (ErrorExpr) exprNode()      {}
func (EmptyExpr) exprNode()      {}
func (ConditionalExpr) exprNode() {}

type BooleanLiteral struct{ Value bool }

func (b BooleanLiteral) String() string {
	if b.Value {
		return "TRUE"
	}
	return "FALSE"
}

type NumberLiteral struct{ Value float64 }

func (n NumberLiteral) String() string {
	return fmt.Sprintf("%v", n.Value)
}

type StringLiteral struct{ Value string }

func (s StringLiteral) String() string {
	return fmt.Sprintf("%q", s.Value)
}

type CellRefExpr struct {
	SheetName string
	Col       string
	Row       int
	ColFixed  bool
	RowFixed  bool
}

func (c CellRefExpr) String() string {
	col := c.Col
	if c.ColFixed {
		col = "$" + col
	}
	row := fmt.Sprintf("%d", c.Row)
	if c.RowFixed {
		row = "$" + row
	}
	if c.SheetName != "" {
		return fmt.Sprintf("%s!%s%s", c.SheetName, col, row)
	}
	return col + row
}

type RangeExpr struct {
	SheetName string
	StartCol  string
	StartRow  int
	EndCol    string
	EndRow    int
}

func (r RangeExpr) String() string {
	if r.SheetName != "" {
		return fmt.Sprintf("%s!%s%d:%s%d", r.SheetName, r.StartCol, r.StartRow, r.EndCol, r.EndRow)
	}
	return fmt.Sprintf("%s%d:%s%d", r.StartCol, r.StartRow, r.EndCol, r.EndRow)
}

type BinaryOpExpr struct {
	Op    string
	Left  Expr
	Right Expr
}

func (b BinaryOpExpr) String() string {
	return fmt.Sprintf("(%s %s %s)", b.Left.String(), b.Op, b.Right.String())
}

type UnaryOpExpr struct {
	Op   string
	Expr Expr
}

func (u UnaryOpExpr) String() string {
	return u.Op + u.Expr.String()
}

type FunctionCallExpr struct {
	Name      string
	Args      []Expr
	NamedArgs map[string]Expr
}

func (f FunctionCallExpr) String() string {
	args := make([]string, len(f.Args))
	for i, a := range f.Args {
		args[i] = a.String()
	}
	return fmt.Sprintf("%s(%s)", f.Name, strings.Join(args, ", "))
}

type ErrorExpr struct{ Code string }

func (e ErrorExpr) String() string { return e.Code }

type EmptyExpr struct{}

func (e EmptyExpr) String() string { return "<empty>" }

type ConditionalExpr struct {
	Condition Expr
	Then      Expr
	Else      Expr
}

func (c ConditionalExpr) String() string {
	return fmt.Sprintf("IF(%s, %s, %s)", c.Condition.String(), c.Then.String(), c.Else.String())
}

func colToIndex(col string) int {
	result := 0
	for _, r := range strings.ToUpper(col) {
		result = result*26 + int(r-'A'+1)
	}
	return result
}

func indexToCol(idx int) string {
	result := ""
	for idx > 0 {
		idx--
		result = string(rune('A'+idx%26)) + result
		idx /= 26
	}
	return result
}

func parseCellRef(ref string) (*CellRefExpr, error) {
	expr := &CellRefExpr{}

	if idx := strings.Index(ref, "!"); idx != -1 {
		expr.SheetName = strings.Trim(ref[:idx], "'")
		ref = ref[idx+1:]
	}

	colFixed := strings.HasPrefix(ref, "$")
	ref = strings.TrimPrefix(ref, "$")
	rowFixed := strings.HasPrefix(ref, "$")
	ref = strings.TrimPrefix(ref, "$")

	col := ""
	for len(ref) > 0 && isLetter(rune(ref[0])) {
		col += string(ref[0])
		ref = ref[1:]
	}

	row := 0
	for len(ref) > 0 && isDigit(rune(ref[0])) {
		row = row*10 + int(ref[0]-'0')
		ref = ref[1:]
	}

	expr.Col = strings.ToUpper(col)
	expr.ColFixed = colFixed
	expr.Row = row
	expr.RowFixed = rowFixed

	return expr, nil
}

type FormulaParser struct {
	tokens []FormulaToken
	pos    int
}

func NewFormulaParser(input string) *FormulaParser {
	tok := NewFormulaTokenizer(input)
	var tokens []FormulaToken
	for {
		t, err := tok.Next()
		if err != nil {
			break
		}
		tokens = append(tokens, t)
		if t.Type == TokenEOF {
			break
		}
	}
	return &FormulaParser{tokens: tokens, pos: 0}
}

func (p *FormulaParser) peek() FormulaToken {
	if p.pos >= len(p.tokens) {
		return FormulaToken{Type: TokenEOF}
	}
	return p.tokens[p.pos]
}

func (p *FormulaParser) next() FormulaToken {
	t := p.peek()
	if p.pos < len(p.tokens) {
		p.pos++
	}
	return t
}

func (p *FormulaParser) Parse() (*FormulaAST, error) {
	e, err := p.parseExpr()
	if err != nil {
		return nil, err
	}
	return &FormulaAST{Root: e}, nil
}

func (p *FormulaParser) parseExpr() (Expr, error) {
	left, err := p.parseTerm()
	if err != nil {
		return nil, err
	}

	for {
		t := p.peek()
		if t.Type == TokenAnd {
			p.next()
			right, err := p.parseTerm()
			if err != nil {
				return nil, err
			}
			left = &BinaryOpExpr{Op: "AND", Left: left, Right: right}
		} else if t.Type == TokenOr {
			p.next()
			right, err := p.parseTerm()
			if err != nil {
				return nil, err
			}
			left = &BinaryOpExpr{Op: "OR", Left: left, Right: right}
		} else if t.Type == TokenAmpersand {
			p.next()
			right, err := p.parseTerm()
			if err != nil {
				return nil, err
			}
			left = &BinaryOpExpr{Op: "||", Left: left, Right: right}
		} else {
			break
		}
	}
	return left, nil
}

func (p *FormulaParser) parseTerm() (Expr, error) {
	left, err := p.parseAdditive()
	if err != nil {
		return nil, err
	}

	for {
		t := p.peek()
		switch t.Type {
		case TokenEQ, TokenNE, TokenLT, TokenGT, TokenLTE, TokenGTE, TokenOperator:
			p.next()
			right, err := p.parseAdditive()
			if err != nil {
				return nil, err
			}
			op := t.Value
			if t.Type == TokenEQ {
				op = "="
			}
			left = &BinaryOpExpr{Op: op, Left: left, Right: right}
		default:
			return left, nil
		}
	}
}

func (p *FormulaParser) parseAdditive() (Expr, error) {
	left, err := p.parseMultiplicative()
	if err != nil {
		return nil, err
	}

	for {
		t := p.peek()
		if t.Type == TokenPlus || t.Type == TokenMinus {
			p.next()
			right, err := p.parseMultiplicative()
			if err != nil {
				return nil, err
			}
			left = &BinaryOpExpr{Op: t.Value, Left: left, Right: right}
		} else {
			break
		}
	}
	return left, nil
}

func (p *FormulaParser) parseMultiplicative() (Expr, error) {
	left, err := p.parseUnary()
	if err != nil {
		return nil, err
	}

	for {
		t := p.peek()
		if t.Type == TokenStar || t.Type == TokenSlash {
			p.next()
			right, err := p.parseUnary()
			if err != nil {
				return nil, err
			}
			left = &BinaryOpExpr{Op: t.Value, Left: left, Right: right}
		} else {
			break
		}
	}
	return left, nil
}

func (p *FormulaParser) parseUnary() (Expr, error) {
	t := p.peek()
	if t.Type == TokenMinus {
		p.next()
		e, err := p.parseUnary()
		if err != nil {
			return nil, err
		}
		return &UnaryOpExpr{Op: "-", Expr: e}, nil
	}
	if t.Type == TokenPlus {
		p.next()
		e, err := p.parseUnary()
		if err != nil {
			return nil, err
		}
		return e, nil
	}
	return p.parsePower()
}

func (p *FormulaParser) parsePower() (Expr, error) {
	left, err := p.parsePostfix()
	if err != nil {
		return nil, err
	}

	t := p.peek()
	if t.Type == TokenCaret {
		p.next()
		right, err := p.parsePower()
		if err != nil {
			return nil, err
		}
		return &BinaryOpExpr{Op: "^", Left: left, Right: right}, nil
	}
	return left, nil
}

func (p *FormulaParser) parsePostfix() (Expr, error) {
	e, err := p.parsePrimary()
	if err != nil {
		return nil, err
	}

	if p.peek().Type == TokenPercent {
		p.next()
		e = &BinaryOpExpr{Op: "/", Left: e, Right: &NumberLiteral{Value: 100}}
	}
	return e, nil
}

func (p *FormulaParser) parsePrimary() (Expr, error) {
	t := p.peek()

	if t.Type == TokenEOF {
		return &EmptyExpr{}, nil
	}

	if t.Type == TokenNumber {
		p.next()
		v := 0.0
		fmt.Sscanf(t.Value, "%f", &v)
		return &NumberLiteral{Value: v}, nil
	}

	if t.Type == TokenString {
		p.next()
		return &StringLiteral{Value: t.Value}, nil
	}

	if t.Type == TokenTRUE {
		p.next()
		return &BooleanLiteral{Value: true}, nil
	}
	if t.Type == TokenFALSE {
		p.next()
		return &BooleanLiteral{Value: false}, nil
	}

	if t.Type == TokenFunction {
		return p.parseFunctionCall()
	}

	if t.Type == TokenCellRef {
		p.next()
		ref := t.Value

		if p.peek().Type == TokenColon {
			p.next()
			endTok := p.peek()
			if endTok.Type == TokenCellRef || endTok.Type == TokenRange {
				var endRef string
				if endTok.Type == TokenCellRef {
					endRef = endTok.Value
				} else {
					endRef = endTok.Value
				}
				p.next()
				start, _ := parseCellRef(ref)
				end, _ := parseCellRef(endRef)
				if start != nil && end != nil {
					return &RangeExpr{
						SheetName: start.SheetName,
						StartCol:  start.Col,
						StartRow:  start.Row,
						EndCol:    end.Col,
						EndRow:    end.Row,
					}, nil
				}
			}
			return &RangeExpr{StartCol: strings.ToUpper(ref), StartRow: 1, EndCol: strings.ToUpper(ref), EndRow: 0}, nil
		}

		cell, err := parseCellRef(ref)
		if err != nil {
			return nil, err
		}
		return cell, nil
	}

	if t.Type == TokenRange {
		p.next()
		parts := strings.Split(t.Value, ":")
		if len(parts) == 2 {
			return &RangeExpr{
				StartCol: strings.ToUpper(parts[0]),
				StartRow: 1,
				EndCol:   strings.ToUpper(parts[1]),
				EndRow:   0,
			}, nil
		}
	}

	if t.Type == TokenLParen {
		p.next()
		e, err := p.parseExpr()
		if err != nil {
			return nil, err
		}
		if p.peek().Type != TokenRParen {
			return nil, fmt.Errorf("expected )")
		}
		p.next()
		return e, nil
	}

	return nil, fmt.Errorf("unexpected token: %s (%v)", t.Value, t.Type)
}

func (p *FormulaParser) parseFunctionCall() (Expr, error) {
	nameTok := p.next().Value
	if p.peek().Type != TokenLParen {
		return nil, fmt.Errorf("expected ( after function name %s", nameTok)
	}
	p.next()

	args := []Expr{}
	namedArgs := make(map[string]Expr)

	if p.peek().Type == TokenRParen {
		p.next()
		return &FunctionCallExpr{Name: nameTok, Args: args, NamedArgs: namedArgs}, nil
	}

	for {
		arg, err := p.parseExpr()
		if err != nil {
			return nil, fmt.Errorf("error parsing argument for %s: %w", nameTok, err)
		}
		// Handle binary operators between arguments (e.g. FILTER(A:C, (B:B="North")*(C:C>100)))
		for {
			t := p.peek()
			if t.Type == TokenComma || t.Type == TokenRParen || t.Type == TokenEOF {
				break
			}
			opTok := t
			p.next()
			right, err := p.parseExpr()
			if err != nil {
				return nil, fmt.Errorf("error parsing argument for %s: %w", nameTok, err)
			}
			op := opTok.Value
			if opTok.Type == TokenEQ {
				op = "="
			}
			arg = &BinaryOpExpr{Op: op, Left: arg, Right: right}
		}
		args = append(args, arg)

		t := p.peek()
		if t.Type == TokenComma {
			p.next()
			continue
		}
		if t.Type == TokenRParen {
			p.next()
			break
		}
		if t.Type == TokenEOF {
			return nil, fmt.Errorf("unterminated function call: %s", nameTok)
		}
		// After a non-comma, non-rparen token (e.g. after ')' in a nested expr),
		// we're at end of function args. Stop parsing more args.
		break
	}

	return &FunctionCallExpr{Name: strings.ToUpper(nameTok), Args: args, NamedArgs: namedArgs}, nil
}

// ============================================================
// SQL GENERATOR
// ============================================================

type SQLGenerator struct {
	tableName string
}

func NewSQLGenerator(tableName string) *SQLGenerator {
	return &SQLGenerator{tableName: tableName}
}

func (g *SQLGenerator) Generate(ast *FormulaAST) (string, error) {
	return g.generateExpr(ast.Root)
}

func (g *SQLGenerator) generateExpr(e Expr) (string, error) {
	switch node := e.(type) {
	case *EmptyExpr:
		return "*", nil
	case *NumberLiteral:
		return fmt.Sprintf("%v", node.Value), nil
	case *StringLiteral:
		return fmt.Sprintf("'%s'", strings.ReplaceAll(node.Value, "'", "''")), nil
	case *BooleanLiteral:
		if node.Value {
			return "1", nil
		}
		return "0", nil
	case *ErrorExpr:
		return "NULL", nil
	case *CellRefExpr:
		col := node.Col
		if node.Row > 0 {
			col = fmt.Sprintf("%s%d", node.Col, node.Row)
		}
		return fmt.Sprintf("\"%s\"", col), nil
	case *RangeExpr:
		return fmt.Sprintf("\"%s\"", node.StartCol), nil
	case *BinaryOpExpr:
		return g.generateBinaryOp(node)
	case *UnaryOpExpr:
		val, err := g.generateExpr(node.Expr)
		if err != nil {
			return "", err
		}
		if node.Op == "-" {
			return fmt.Sprintf("-%s", val), nil
		}
		return val, nil
	case *FunctionCallExpr:
		return g.generateFunction(node)
	case *ConditionalExpr:
		cond, err := g.generateExpr(node.Condition)
		if err != nil {
			return "", err
		}
		then, err := g.generateExpr(node.Then)
		if err != nil {
			return "", err
		}
		elseVal := "NULL"
		if node.Else != nil {
			var err error
			elseVal, err = g.generateExpr(node.Else)
			if err != nil {
				return "", err
			}
		}
		return fmt.Sprintf("CASE WHEN %s THEN %s ELSE %s END", cond, then, elseVal), nil
	default:
		return "", fmt.Errorf("unsupported expression type: %T", e)
	}
}

func (g *SQLGenerator) generateBinaryOp(node *BinaryOpExpr) (string, error) {
	left, err := g.generateExpr(node.Left)
	if err != nil {
		return "", err
	}
	right, err := g.generateExpr(node.Right)
	if err != nil {
		return "", err
	}
	switch node.Op {
	case "+", "-", "*", "/", "^":
		return fmt.Sprintf("(%s %s %s)", left, node.Op, right), nil
	case "=":
		return fmt.Sprintf("(%s = %s)", left, right), nil
	case "<>":
		return fmt.Sprintf("(%s <> %s)", left, right), nil
	case "<":
		return fmt.Sprintf("(%s < %s)", left, right), nil
	case ">":
		return fmt.Sprintf("(%s > %s)", left, right), nil
	case "<=":
		return fmt.Sprintf("(%s <= %s)", left, right), nil
	case ">=":
		return fmt.Sprintf("(%s >= %s)", left, right), nil
	case "||":
		return fmt.Sprintf("(%s || %s)", left, right), nil
	case "AND":
		return fmt.Sprintf("(%s AND %s)", left, right), nil
	case "OR":
		return fmt.Sprintf("(%s OR %s)", left, right), nil
	default:
		return "", fmt.Errorf("unsupported operator: %s", node.Op)
	}
}

func (g *SQLGenerator) generateFunction(node *FunctionCallExpr) (string, error) {
	name := node.Name
	args := node.Args

	switch name {
	case "SUM":
		if len(args) == 0 {
			return "", fmt.Errorf("SUM requires at least one argument")
		}
		expr, _ := g.generateExpr(args[0])
		return fmt.Sprintf("SUM(%s)", expr), nil
	case "AVERAGE", "AVG":
		if len(args) == 0 {
			return "", fmt.Errorf("AVERAGE requires at least one argument")
		}
		expr, _ := g.generateExpr(args[0])
		return fmt.Sprintf("AVG(%s)", expr), nil
	case "COUNT", "COUNTA":
		if len(args) == 0 {
			return "COUNT(*)", nil
		}
		expr, _ := g.generateExpr(args[0])
		return fmt.Sprintf("COUNT(%s)", expr), nil
	case "COUNTBLANK":
		if len(args) == 0 {
			return "COUNT(*) - COUNT(*)", nil
		}
		expr, _ := g.generateExpr(args[0])
		return fmt.Sprintf("COUNT(CASE WHEN %s IS NULL OR %s = '' THEN 1 END)", expr, expr), nil
	case "MAX":
		if len(args) == 0 {
			return "", fmt.Errorf("MAX requires at least one argument")
		}
		expr, _ := g.generateExpr(args[0])
		return fmt.Sprintf("MAX(%s)", expr), nil
	case "MIN":
		if len(args) == 0 {
			return "", fmt.Errorf("MIN requires at least one argument")
		}
		expr, _ := g.generateExpr(args[0])
		return fmt.Sprintf("MIN(%s)", expr), nil
	case "SUMIF":
		return g.genSUMIF(node)
	case "SUMIFS":
		return g.genSUMIFS(node)
	case "COUNTIF":
		return g.genCOUNTIF(node)
	case "COUNTIFS":
		return g.genCOUNTIFS(node)
	case "AVERAGEIF":
		return g.genAVERAGEIF(node)
	case "MAXIFS":
		return g.genMAXIFS(node)
	case "MINIFS":
		return g.genMINIFS(node)
	case "VLOOKUP":
		return g.genVLOOKUP(node)
	case "INDEX":
		return g.genINDEX(node)
	case "MATCH":
		return g.genMATCH(node)
	case "XLOOKUP":
		return g.genXLOOKUP(node)
	case "HLOOKUP":
		return g.genVLOOKUP(node)
	case "LOOKUP":
		return g.genLOOKUP(node)
	case "IF":
		return g.genIF(node)
	case "IFERROR", "IFNA", "ISERROR", "ISERR", "ISNA":
		return g.genIFERROR(node)
	case "AND":
		return g.genAND(node)
	case "OR":
		return g.genOR(node)
	case "NOT":
		if len(args) == 0 {
			return "", fmt.Errorf("NOT requires one argument")
		}
		expr, _ := g.generateExpr(args[0])
		return fmt.Sprintf("NOT %s", expr), nil
	case "IFS":
		return g.genIFS(node)
	case "SWITCH":
		return g.genSWITCH(node)
	case "LEFT":
		return g.genLEFT(node)
	case "RIGHT":
		return g.genRIGHT(node)
	case "MID":
		return g.genMID(node)
	case "LEN", "LENGTH":
		if len(args) == 0 {
			return "", fmt.Errorf("LEN requires one argument")
		}
		expr, _ := g.generateExpr(args[0])
		return fmt.Sprintf("LENGTH(%s)", expr), nil
	case "UPPER":
		if len(args) == 0 {
			return "", fmt.Errorf("UPPER requires one argument")
		}
		expr, _ := g.generateExpr(args[0])
		return fmt.Sprintf("UPPER(%s)", expr), nil
	case "LOWER":
		if len(args) == 0 {
			return "", fmt.Errorf("LOWER requires one argument")
		}
		expr, _ := g.generateExpr(args[0])
		return fmt.Sprintf("LOWER(%s)", expr), nil
	case "TRIM":
		if len(args) == 0 {
			return "", fmt.Errorf("TRIM requires one argument")
		}
		expr, _ := g.generateExpr(args[0])
		return fmt.Sprintf("TRIM(%s)", expr), nil
	case "CONCAT", "CONCATENATE":
		return g.genCONCAT(node)
	case "TEXT":
		if len(args) == 0 {
			return "", fmt.Errorf("TEXT requires at least one argument")
		}
		expr, _ := g.generateExpr(args[0])
		return fmt.Sprintf("CAST(%s AS TEXT)", expr), nil
	case "SUBSTITUTE":
		if len(args) < 3 {
			return "", fmt.Errorf("SUBSTITUTE requires at least 3 arguments")
		}
		expr, _ := g.generateExpr(args[0])
		old, _ := g.generateExpr(args[1])
		new, _ := g.generateExpr(args[2])
		return fmt.Sprintf("REPLACE(%s, %s, %s)", expr, old, new), nil
	case "FIND", "SEARCH":
		if len(args) < 2 {
			return "", fmt.Errorf("%s requires at least 2 arguments", name)
		}
		needle, _ := g.generateExpr(args[0])
		haystack, _ := g.generateExpr(args[1])
		if name == "FIND" {
			return fmt.Sprintf("INSTR(%s, %s)", haystack, needle), nil
		}
		return fmt.Sprintf("INSTR(LOWER(%s), LOWER(%s))", haystack, needle), nil
	case "REPT":
		if len(args) < 2 {
			return "", fmt.Errorf("REPT requires 2 arguments")
		}
		expr, _ := g.generateExpr(args[0])
		n, _ := g.generateExpr(args[1])
		return fmt.Sprintf("REPLACE(PRINTF('%%-%s', ''), ' ', %s)", n, expr), nil
	case "VALUE":
		if len(args) == 0 {
			return "", fmt.Errorf("VALUE requires one argument")
		}
		expr, _ := g.generateExpr(args[0])
		return fmt.Sprintf("CAST(%s AS NUMERIC)", expr), nil
	case "TEXTJOIN":
		if len(args) < 3 {
			return "", fmt.Errorf("TEXTJOIN requires at least 3 arguments")
		}
		delim, _ := g.generateExpr(args[0])
		parts := make([]string, 0)
		for i := 2; i < len(args); i++ {
			p, _ := g.generateExpr(args[i])
			parts = append(parts, p)
		}
		return fmt.Sprintf("(%s)", strings.Join(parts, " || "+delim+" || ")), nil
	case "ABS":
		if len(args) == 0 {
			return "", fmt.Errorf("ABS requires one argument")
		}
		expr, _ := g.generateExpr(args[0])
		return fmt.Sprintf("ABS(%s)", expr), nil
	case "ROUND":
		expr := "0"
		if len(args) >= 1 {
			expr, _ = g.generateExpr(args[0])
		}
		digits := "0"
		if len(args) >= 2 {
			digits, _ = g.generateExpr(args[1])
		}
		return fmt.Sprintf("ROUND(%s, CAST(%s AS INTEGER))", expr, digits), nil
	case "INT":
		if len(args) == 0 {
			return "", fmt.Errorf("INT requires one argument")
		}
		expr, _ := g.generateExpr(args[0])
		return fmt.Sprintf("CAST(%s AS INTEGER)", expr), nil
	case "SQRT":
		if len(args) == 0 {
			return "", fmt.Errorf("SQRT requires one argument")
		}
		expr, _ := g.generateExpr(args[0])
		return fmt.Sprintf("SQRT(%s)", expr), nil
	case "MOD":
		if len(args) < 2 {
			return "", fmt.Errorf("MOD requires 2 arguments")
		}
		a, _ := g.generateExpr(args[0])
		b, _ := g.generateExpr(args[1])
		return fmt.Sprintf("(%s %% %s)", a, b), nil
	case "POWER":
		if len(args) < 2 {
			return "", fmt.Errorf("POWER requires 2 arguments")
		}
		a, _ := g.generateExpr(args[0])
		b, _ := g.generateExpr(args[1])
		return fmt.Sprintf("POWER(%s, %s)", a, b), nil
	case "RAND":
		return "RANDOM()", nil
	case "RANDBETWEEN":
		if len(args) < 2 {
			return "", fmt.Errorf("RANDBETWEEN requires 2 arguments")
		}
		low, _ := g.generateExpr(args[0])
		high, _ := g.generateExpr(args[1])
		return fmt.Sprintf("(ABS(RANDOM()) %% (CAST(%s AS INTEGER) - CAST(%s AS INTEGER) + 1) + CAST(%s AS INTEGER))", high, low, low), nil
	case "SIGN":
		if len(args) == 0 {
			return "", fmt.Errorf("SIGN requires one argument")
		}
		expr, _ := g.generateExpr(args[0])
		return fmt.Sprintf("CASE WHEN %s > 0 THEN 1 WHEN %s < 0 THEN -1 ELSE 0 END", expr, expr), nil
	case "LN":
		if len(args) == 0 {
			return "", fmt.Errorf("LN requires one argument")
		}
		expr, _ := g.generateExpr(args[0])
		return fmt.Sprintf("LN(%s)", expr), nil
	case "LOG10":
		if len(args) == 0 {
			return "", fmt.Errorf("LOG10 requires one argument")
		}
		expr, _ := g.generateExpr(args[0])
		return fmt.Sprintf("LOG10(%s)", expr), nil
	case "LOG":
		if len(args) == 0 {
			return "", fmt.Errorf("LOG requires one argument")
		}
		expr, _ := g.generateExpr(args[0])
		if len(args) >= 2 {
			base, _ := g.generateExpr(args[1])
			return fmt.Sprintf("(LOG(%s) / LOG(%s))", expr, base), nil
		}
		return fmt.Sprintf("LOG(%s)", expr), nil
	case "EXP":
		if len(args) == 0 {
			return "", fmt.Errorf("EXP requires one argument")
		}
		expr, _ := g.generateExpr(args[0])
		return fmt.Sprintf("EXP(%s)", expr), nil
	case "TODAY":
		return "date('now')", nil
	case "NOW":
		return "datetime('now')", nil
	case "DATE":
		if len(args) < 3 {
			return "", fmt.Errorf("DATE requires 3 arguments")
		}
		y, _ := g.generateExpr(args[0])
		m, _ := g.generateExpr(args[1])
		d, _ := g.generateExpr(args[2])
		return fmt.Sprintf("DATE(CAST(%s AS TEXT) || '-' || PRINTF('%%02d', CAST(%s AS INTEGER)) || '-' || PRINTF('%%02d', CAST(%s AS INTEGER)))", y, m, d), nil
	case "YEAR":
		if len(args) == 0 {
			return "", fmt.Errorf("YEAR requires one argument")
		}
		expr, _ := g.generateExpr(args[0])
		return fmt.Sprintf("CAST(strftime('%%Y', %s) AS INTEGER)", expr), nil
	case "MONTH":
		if len(args) == 0 {
			return "", fmt.Errorf("MONTH requires one argument")
		}
		expr, _ := g.generateExpr(args[0])
		return fmt.Sprintf("CAST(strftime('%%m', %s) AS INTEGER)", expr), nil
	case "DAY":
		if len(args) == 0 {
			return "", fmt.Errorf("DAY requires one argument")
		}
		expr, _ := g.generateExpr(args[0])
		return fmt.Sprintf("CAST(strftime('%%d', %s) AS INTEGER)", expr), nil
	case "DATEDIF":
		if len(args) < 3 {
			return "", fmt.Errorf("DATEDIF requires 3 arguments")
		}
		start, _ := g.generateExpr(args[0])
		end, _ := g.generateExpr(args[1])
		unit, _ := g.generateExpr(args[2])
		unitStr := strings.Trim(unit, "'\"")
		switch unitStr {
		case "D":
			return fmt.Sprintf("CAST((julianday(%s) - julianday(%s)) AS INTEGER)", end, start), nil
		case "M":
			return fmt.Sprintf("CAST((strftime('%%Y', %s) - strftime('%%Y', %s)) * 12 + (strftime('%%m', %s) - strftime('%%m', %s)) AS INTEGER)", end, start, end, start), nil
		case "Y":
			return fmt.Sprintf("CAST(strftime('%%Y', %s) - strftime('%%Y', %s) AS INTEGER)", end, start), nil
		default:
			return fmt.Sprintf("CAST((julianday(%s) - julianday(%s)) AS INTEGER)", end, start), nil
		}
	case "EDATE":
		if len(args) < 2 {
			return "", fmt.Errorf("EDATE requires 2 arguments")
		}
		start, _ := g.generateExpr(args[0])
		months, _ := g.generateExpr(args[1])
		return fmt.Sprintf("DATE(%s, 'start of month', '+' || CAST(%s AS INTEGER) || ' months', '+1 month', '-1 day')", start, months), nil
	case "EOMONTH":
		if len(args) < 2 {
			return "", fmt.Errorf("EOMONTH requires 2 arguments")
		}
		start, _ := g.generateExpr(args[0])
		months, _ := g.generateExpr(args[1])
		return fmt.Sprintf("DATE(%s, 'start of month', '+' || (CAST(%s AS INTEGER) + 1) || ' months', '-1 day')", start, months), nil
	case "WEEKDAY":
		if len(args) == 0 {
			return "", fmt.Errorf("WEEKDAY requires one argument")
		}
		expr, _ := g.generateExpr(args[0])
		return fmt.Sprintf("CAST(strftime('%%w', %s) + 1 AS INTEGER)", expr), nil
	case "NA":
		return "NULL", nil
	case "FILTER":
		return g.genFILTER(node)
	case "SORT":
		return g.genSORT(node)
	case "UNIQUE":
		return g.genUNIQUE(node)
	case "LET":
		if len(node.Args) < 3 {
			return "", fmt.Errorf("LET requires at least 3 arguments")
		}
		last, _ := g.generateExpr(node.Args[len(node.Args)-1])
		return last, nil
	case "SEQUENCE":
		return "", fmt.Errorf("SEQUENCE is not supported in SQL conversion")
	case "VSTACK", "HSTACK":
		return "", fmt.Errorf("%s is not supported in SQL conversion", name)
	case "BYROW", "BYCOL":
		return "", fmt.Errorf("%s is not supported in SQL conversion", name)
	case "MAP":
		return "", fmt.Errorf("MAP is not supported in SQL conversion")
	default:
		return "", fmt.Errorf("unsupported function: %s", name)
	}
}

// ---- Aggregate generators ----

func (g *SQLGenerator) genSUMIF(node *FunctionCallExpr) (string, error) {
	if len(node.Args) < 2 {
		return "", fmt.Errorf("SUMIF requires at least 2 arguments")
	}
	sumCol, _ := g.generateExpr(node.Args[0])
	crit, _ := g.generateExpr(node.Args[1])
	sumExpr := sumCol
	if len(node.Args) >= 3 {
		sumExpr, _ = g.generateExpr(node.Args[2])
	}
	return fmt.Sprintf("SUM(CASE WHEN %s = %s THEN %s ELSE 0 END)", sumCol, crit, sumExpr), nil
}

func (g *SQLGenerator) genSUMIFS(node *FunctionCallExpr) (string, error) {
	if len(node.Args) < 3 {
		return "", fmt.Errorf("SUMIFS requires at least 3 arguments")
	}
	sumExpr, _ := g.generateExpr(node.Args[0])
	conditions := []string{}
	for i := 1; i+1 < len(node.Args); i += 2 {
		critRange, _ := g.generateExpr(node.Args[i])
		crit, _ := g.generateExpr(node.Args[i+1])
		conditions = append(conditions, fmt.Sprintf("%s = %s", critRange, crit))
	}
	return fmt.Sprintf("SUM(CASE WHEN %s THEN %s ELSE 0 END)", strings.Join(conditions, " AND "), sumExpr), nil
}

func (g *SQLGenerator) genCOUNTIF(node *FunctionCallExpr) (string, error) {
	if len(node.Args) < 2 {
		return "", fmt.Errorf("COUNTIF requires 2 arguments")
	}
	col, _ := g.generateExpr(node.Args[0])
	crit, _ := g.generateExpr(node.Args[1])
	return fmt.Sprintf("COUNT(CASE WHEN %s = %s THEN 1 END)", col, crit), nil
}

func (g *SQLGenerator) genCOUNTIFS(node *FunctionCallExpr) (string, error) {
	if len(node.Args) < 3 {
		return "", fmt.Errorf("COUNTIFS requires at least 3 arguments")
	}
	conditions := []string{}
	for i := 0; i+1 < len(node.Args); i += 2 {
		col, _ := g.generateExpr(node.Args[i])
		crit, _ := g.generateExpr(node.Args[i+1])
		conditions = append(conditions, fmt.Sprintf("%s = %s", col, crit))
	}
	return fmt.Sprintf("COUNT(CASE WHEN %s THEN 1 END)", strings.Join(conditions, " AND ")), nil
}

func (g *SQLGenerator) genAVERAGEIF(node *FunctionCallExpr) (string, error) {
	if len(node.Args) < 2 {
		return "", fmt.Errorf("AVERAGEIF requires at least 2 arguments")
	}
	col, _ := g.generateExpr(node.Args[0])
	crit, _ := g.generateExpr(node.Args[1])
	avgCol := col
	if len(node.Args) >= 3 {
		avgCol, _ = g.generateExpr(node.Args[2])
	}
	return fmt.Sprintf("AVG(CASE WHEN %s = %s THEN %s END)", col, crit, avgCol), nil
}

func (g *SQLGenerator) genMAXIFS(node *FunctionCallExpr) (string, error) {
	if len(node.Args) < 4 {
		return "", fmt.Errorf("MAXIFS requires at least 4 arguments")
	}
	maxCol, _ := g.generateExpr(node.Args[0])
	conditions := []string{}
	for i := 1; i+1 < len(node.Args); i += 2 {
		c, _ := g.generateExpr(node.Args[i])
		v, _ := g.generateExpr(node.Args[i+1])
		conditions = append(conditions, fmt.Sprintf("%s = %s", c, v))
	}
	return fmt.Sprintf("MAX(CASE WHEN %s THEN %s END)", strings.Join(conditions, " AND "), maxCol), nil
}

func (g *SQLGenerator) genMINIFS(node *FunctionCallExpr) (string, error) {
	if len(node.Args) < 4 {
		return "", fmt.Errorf("MINIFS requires at least 4 arguments")
	}
	minCol, _ := g.generateExpr(node.Args[0])
	conditions := []string{}
	for i := 1; i+1 < len(node.Args); i += 2 {
		c, _ := g.generateExpr(node.Args[i])
		v, _ := g.generateExpr(node.Args[i+1])
		conditions = append(conditions, fmt.Sprintf("%s = %s", c, v))
	}
	return fmt.Sprintf("MIN(CASE WHEN %s THEN %s END)", strings.Join(conditions, " AND "), minCol), nil
}

// ---- Lookup generators ----

func (g *SQLGenerator) genVLOOKUP(node *FunctionCallExpr) (string, error) {
	if len(node.Args) < 3 {
		return "", fmt.Errorf("VLOOKUP requires at least 3 arguments")
	}
	lookupVal, _ := g.generateExpr(node.Args[0])
	colIdx := 1
	if len(node.Args) >= 3 {
		if lit, ok := node.Args[2].(NumberLiteral); ok {
			colIdx = int(lit.Value)
		}
	}
	exactMatch := true
	if len(node.Args) >= 4 {
		if lit, ok := node.Args[3].(NumberLiteral); ok && lit.Value != 0 {
			exactMatch = false
		}
	}
	lookupCol := "A"
	returnCol := indexToCol(colIdx)
	if exactMatch {
		return fmt.Sprintf("(SELECT \"%s\" FROM \"%s\" WHERE \"%s\" = %s LIMIT 1)", returnCol, g.tableName, lookupCol, lookupVal), nil
	}
	return fmt.Sprintf("(SELECT \"%s\" FROM \"%s\" WHERE \"%s\" <= %s ORDER BY \"%s\" DESC LIMIT 1)", returnCol, g.tableName, lookupCol, lookupVal, lookupCol), nil
}

func (g *SQLGenerator) genINDEX(node *FunctionCallExpr) (string, error) {
	if len(node.Args) < 2 {
		return "", fmt.Errorf("INDEX requires at least 2 arguments")
	}
	arrayExpr, _ := g.generateExpr(node.Args[0])
	rowNum, _ := g.generateExpr(node.Args[1])
	return fmt.Sprintf("(SELECT \"%s\" FROM (SELECT ROW_NUMBER() OVER() - 1 AS __rn, \"%s\" FROM \"%s\") t WHERE __rn = %s - 1 LIMIT 1)", arrayExpr, arrayExpr, g.tableName, rowNum), nil
}

func (g *SQLGenerator) genMATCH(node *FunctionCallExpr) (string, error) {
	if len(node.Args) < 2 {
		return "", fmt.Errorf("MATCH requires at least 2 arguments")
	}
	lookupVal, _ := g.generateExpr(node.Args[0])
	lookupArray, _ := g.generateExpr(node.Args[1])
	matchType := "0"
	if len(node.Args) >= 3 {
		matchType, _ = g.generateExpr(node.Args[2])
	}
	exact := matchType == "0"
	if exact {
		return fmt.Sprintf("(SELECT ROWID FROM \"%s\" WHERE \"%s\" = %s LIMIT 1) - 1", g.tableName, lookupArray, lookupVal), nil
	}
	return fmt.Sprintf("(SELECT ROWID FROM \"%s\" WHERE \"%s\" <= %s ORDER BY \"%s\" DESC LIMIT 1) - 1", g.tableName, lookupArray, lookupVal, lookupArray), nil
}

func (g *SQLGenerator) genXLOOKUP(node *FunctionCallExpr) (string, error) {
	if len(node.Args) < 3 {
		return "", fmt.Errorf("XLOOKUP requires at least 3 arguments")
	}
	lookupVal, _ := g.generateExpr(node.Args[0])
	lookupArray, _ := g.generateExpr(node.Args[1])
	returnArray, _ := g.generateExpr(node.Args[2])
	ifNotFound := "NULL"
	if len(node.Args) >= 4 {
		ifNotFound, _ = g.generateExpr(node.Args[3])
	}
	return fmt.Sprintf("(SELECT COALESCE(\"%s\", %s) FROM \"%s\" WHERE \"%s\" = %s LIMIT 1)", returnArray, ifNotFound, g.tableName, lookupArray, lookupVal), nil
}

func (g *SQLGenerator) genLOOKUP(node *FunctionCallExpr) (string, error) {
	if len(node.Args) < 2 {
		return "", fmt.Errorf("LOOKUP requires at least 2 arguments")
	}
	lookupVal, _ := g.generateExpr(node.Args[0])
	lookupArray, _ := g.generateExpr(node.Args[1])
	returnArray := lookupArray
	if len(node.Args) >= 3 {
		returnArray, _ = g.generateExpr(node.Args[2])
	}
	return fmt.Sprintf("(SELECT \"%s\" FROM \"%s\" WHERE \"%s\" <= %s ORDER BY \"%s\" LIMIT 1)", returnArray, g.tableName, lookupArray, lookupVal, lookupArray), nil
}

// ---- Logical generators ----

func (g *SQLGenerator) genIF(node *FunctionCallExpr) (string, error) {
	if len(node.Args) < 2 {
		return "", fmt.Errorf("IF requires at least 2 arguments")
	}
	cond, _ := g.generateExpr(node.Args[0])
	then, _ := g.generateExpr(node.Args[1])
	elseVal := "NULL"
	if len(node.Args) >= 3 {
		elseVal, _ = g.generateExpr(node.Args[2])
	}
	return fmt.Sprintf("CASE WHEN %s THEN %s ELSE %s END", cond, then, elseVal), nil
}

func (g *SQLGenerator) genIFERROR(node *FunctionCallExpr) (string, error) {
	if len(node.Args) < 2 {
		return "", fmt.Errorf("IFERROR requires 2 arguments")
	}
	expr, _ := g.generateExpr(node.Args[0])
	fallback, _ := g.generateExpr(node.Args[1])
	return fmt.Sprintf("COALESCE(NULLIF(%s, ''), %s)", expr, fallback), nil
}

func (g *SQLGenerator) genAND(node *FunctionCallExpr) (string, error) {
	if len(node.Args) == 0 {
		return "", fmt.Errorf("AND requires at least one argument")
	}
	parts := make([]string, len(node.Args))
	for i, arg := range node.Args {
		p, _ := g.generateExpr(arg)
		parts[i] = p
	}
	return strings.Join(parts, " AND "), nil
}

func (g *SQLGenerator) genOR(node *FunctionCallExpr) (string, error) {
	if len(node.Args) == 0 {
		return "", fmt.Errorf("OR requires at least one argument")
	}
	parts := make([]string, len(node.Args))
	for i, arg := range node.Args {
		p, _ := g.generateExpr(arg)
		parts[i] = p
	}
	return strings.Join(parts, " OR "), nil
}

func (g *SQLGenerator) genIFS(node *FunctionCallExpr) (string, error) {
	if len(node.Args) < 2 {
		return "", fmt.Errorf("IFS requires at least 2 arguments")
	}
	var parts []string
	for i := 0; i+1 < len(node.Args); i += 2 {
		cond, _ := g.generateExpr(node.Args[i])
		val, _ := g.generateExpr(node.Args[i+1])
		parts = append(parts, fmt.Sprintf("WHEN %s THEN %s", cond, val))
	}
	return "CASE " + strings.Join(parts, " ") + " END", nil
}

func (g *SQLGenerator) genSWITCH(node *FunctionCallExpr) (string, error) {
	if len(node.Args) < 3 {
		return "", fmt.Errorf("SWITCH requires at least 3 arguments")
	}
	expr, _ := g.generateExpr(node.Args[0])
	var parts []string
	for i := 1; i+1 < len(node.Args); i += 2 {
		m, _ := g.generateExpr(node.Args[i])
		v, _ := g.generateExpr(node.Args[i+1])
		parts = append(parts, fmt.Sprintf("WHEN %s = %s THEN %s", expr, m, v))
	}
	if len(node.Args)%2 == 1 {
		defaultVal, _ := g.generateExpr(node.Args[len(node.Args)-1])
		parts = append(parts, fmt.Sprintf("ELSE %s", defaultVal))
	}
	return "CASE " + strings.Join(parts, " ") + " END", nil
}

// ---- Text generators ----

func (g *SQLGenerator) genLEFT(node *FunctionCallExpr) (string, error) {
	if len(node.Args) < 1 {
		return "", fmt.Errorf("LEFT requires at least one argument")
	}
	col, _ := g.generateExpr(node.Args[0])
	n := "255"
	if len(node.Args) >= 2 {
		n, _ = g.generateExpr(node.Args[1])
	}
	return fmt.Sprintf("SUBSTR(%s, 1, %s)", col, n), nil
}

func (g *SQLGenerator) genRIGHT(node *FunctionCallExpr) (string, error) {
	if len(node.Args) < 1 {
		return "", fmt.Errorf("RIGHT requires at least one argument")
	}
	col, _ := g.generateExpr(node.Args[0])
	n := "1"
	if len(node.Args) >= 2 {
		n, _ = g.generateExpr(node.Args[1])
	}
	return fmt.Sprintf("SUBSTR(%s, -CAST(%s AS INTEGER), %s)", col, n, n), nil
}

func (g *SQLGenerator) genMID(node *FunctionCallExpr) (string, error) {
	if len(node.Args) < 3 {
		return "", fmt.Errorf("MID requires 3 arguments")
	}
	col, _ := g.generateExpr(node.Args[0])
	start, _ := g.generateExpr(node.Args[1])
	n, _ := g.generateExpr(node.Args[2])
	return fmt.Sprintf("SUBSTR(%s, CAST(%s AS INTEGER), CAST(%s AS INTEGER))", col, start, n), nil
}

func (g *SQLGenerator) genCONCAT(node *FunctionCallExpr) (string, error) {
	parts := make([]string, len(node.Args))
	for i, arg := range node.Args {
		p, _ := g.generateExpr(arg)
		parts[i] = p
	}
	return strings.Join(parts, " || "), nil
}

// ---- Array generators ----

func (g *SQLGenerator) genFILTER(node *FunctionCallExpr) (string, error) {
	if len(node.Args) < 2 {
		return "", fmt.Errorf("FILTER requires at least 2 arguments")
	}
	include, _ := g.generateExpr(node.Args[1])
	return fmt.Sprintf("(SELECT * FROM \"%s\" WHERE %s)", g.tableName, include), nil
}

func (g *SQLGenerator) genSORT(node *FunctionCallExpr) (string, error) {
	if len(node.Args) < 2 {
		return "", fmt.Errorf("SORT requires at least 2 arguments")
	}
	orderCol, _ := g.generateExpr(node.Args[1])
	orderDir := "ASC"
	if len(node.Args) >= 3 {
		dir, _ := g.generateExpr(node.Args[2])
		if dir == "-1" || dir == "'-1'" {
			orderDir = "DESC"
		}
	}
	return fmt.Sprintf("(SELECT * FROM \"%s\" ORDER BY %s %s)", g.tableName, orderCol, orderDir), nil
}

func (g *SQLGenerator) genUNIQUE(node *FunctionCallExpr) (string, error) {
	if len(node.Args) == 0 {
		return "", fmt.Errorf("UNIQUE requires at least one argument")
	}
	col, _ := g.generateExpr(node.Args[0])
	return fmt.Sprintf("SELECT DISTINCT %s FROM \"%s\"", col, g.tableName), nil
}
