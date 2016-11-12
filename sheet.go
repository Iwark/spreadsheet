package spreadsheet

import "encoding/json"

// Sheet is a sheet in a spreadsheet.
type Sheet struct {
	Properties SheetProperties `json:"properties"`
	Data       SheetData       `json:"data"`
	// Merges []*GridRange `json:"merges"`
	// ConditionalFormats []*ConditionalFormatRule `json:"conditionalFormats"`
	// FilterViews []*FilterView `json:"filterViews"`
	// ProtectedRanges []*ProtectedRange `json:"protectedRanges"`
	// BasicFilter *BasicFilter `json:"basicFilter"`
	// Charts []*EmbeddedChart `json:"charts"`
	// BandedRanges []*BandedRange `json:"bandedRanges"`

	Spreadsheet *Spreadsheet `json:"-"`
	Rows        [][]Cell     `json:"-"`
	Columns     [][]Cell     `json:"-"`

	modifiedCells []*Cell
	newMaxRow     uint
	newMaxColumn  uint
}

// UnmarshalJSON embeds rows and columns to the sheet.
func (sheet *Sheet) UnmarshalJSON(data []byte) error {
	type Alias Sheet
	a := (*Alias)(sheet)
	if err := json.Unmarshal(data, a); err != nil {
		return err
	}
	var maxRow, maxColumn int
	cells := []Cell{}
	for _, gridData := range sheet.Data.GridData {
		for rowNum, row := range gridData.RowData {
			for columnNum, cellData := range row.Values {
				r := gridData.StartRow + uint(rowNum)
				if int(r) > maxRow {
					maxRow = int(r)
				}
				c := gridData.StartColumn + uint(columnNum)
				if int(c) > maxColumn {
					maxColumn = int(c)
				}
				cell := Cell{
					Row:    r,
					Column: c,
					Value:  cellData.FormattedValue,
				}
				cells = append(cells, cell)
			}
		}
	}
	sheet.Rows = make([][]Cell, maxRow+1)
	for i := range sheet.Rows {
		sheet.Rows[i] = make([]Cell, maxColumn+1)
	}
	sheet.Columns = make([][]Cell, maxColumn+1)
	for i := range sheet.Columns {
		sheet.Columns[i] = make([]Cell, maxRow+1)
	}
	for _, cell := range cells {
		sheet.Rows[cell.Row][cell.Column] = cell
		sheet.Columns[cell.Column][cell.Row] = cell
	}

	sheet.modifiedCells = []*Cell{}
	sheet.newMaxRow = sheet.Properties.GridProperties.RowCount
	sheet.newMaxColumn = sheet.Properties.GridProperties.ColumnCount

	return nil
}

// Update updates cell changes
func (sheet *Sheet) Update(row, column int, val string) {
	if uint(row)+1 > sheet.newMaxRow {
		sheet.newMaxRow = uint(row) + 1
	}
	if uint(column)+1 > sheet.newMaxColumn {
		sheet.newMaxColumn = uint(column) + 1
	}
	for _, cell := range sheet.modifiedCells {
		if cell.Row == uint(row) && cell.Column == uint(column) {
			cell.Value = val
			return
		}
	}
	sheet.modifiedCells = append(sheet.modifiedCells, &Cell{
		Row:    uint(row),
		Column: uint(column),
		Value:  val,
	})
}

// Synchronize reflects the changes of the sheet.
func (sheet *Sheet) Synchronize() (err error) {
	err = sheet.Spreadsheet.service.SyncSheet(sheet)
	return
}
