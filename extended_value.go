package spreadsheet

// ExtendedValue is the kinds of value that a cell in a spreadsheet can have.
type ExtendedValue struct {
	NumberValue  float64    `json:"numberValue"`
	StringValue  string     `json:"stringValue"`
	BoolValue    bool       `json:"boolValue"`
	FormulaValue string     `json:"formulaValue"`
	ErrorValue   ErrorValue `json:"errorValue"`
}
