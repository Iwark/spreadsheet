package spreadsheet

// GridProperties is properties of a grid.
type GridProperties struct {
	RowCount          uint `json:"rowCount"`
	ColumnCount       uint `json:"columnCount"`
	FrozenRowCount    uint `json:"frozenRowCount"`
	FrozenColumnCount uint `json:"frozenColumnCount"`
	HideGridlines     bool `json:"hideGridlines"`
}
