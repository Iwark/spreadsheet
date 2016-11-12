package spreadsheet

// ErrorValue is an error in a cell.
type ErrorValue struct {
	Type    string `json:"type"`
	Message string `json:"message"`
}
