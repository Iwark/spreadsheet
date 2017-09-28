package spreadsheet

// GridData is data in the grid, as well as metadata about the dimensions.
type GridData struct {
	StartRow       uint                   `json:"startRow"`
	StartColumn    uint                   `json:"startColumn"`
	RowData        []RowData              `json:"rowData"`
	RowMetadata    []*DimensionProperties `json:"rowMetadata"`
	ColumnMetadata []*DimensionProperties `json:"columnMetadata"`
}
