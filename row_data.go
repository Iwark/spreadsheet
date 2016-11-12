package spreadsheet

// RowData is data about each cell in a row.
type RowData struct {
	Values []CellData `json:"values"`
}
