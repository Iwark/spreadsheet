package spreadsheet

import (
	"encoding/json"
)

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
	sheet.Rows, sheet.Columns = newCells(uint(maxRow), uint(maxColumn))

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

	if uint(len(sheet.Rows)) < sheet.newMaxRow+1 ||
		uint(len(sheet.Columns)) < sheet.newMaxColumn+1 {
		sheet.Rows = appendCells(sheet.Rows, sheet.newMaxRow, sheet.newMaxColumn, func(i, t uint) Cell {
			return Cell{Row: i, Column: t}
		})
		sheet.Columns = appendCells(sheet.Columns, sheet.newMaxColumn, sheet.newMaxRow, func(i, t uint) Cell {
			return Cell{Row: t, Column: i}
		})
	}

	cell := Cell{
		Row:    uint(row),
		Column: uint(column),
		Value:  val,
	}

	sheet.Rows[row][column] = cell
	sheet.Columns[column][row] = cell
	for _, cell := range sheet.modifiedCells {
		if cell.Row == uint(row) && cell.Column == uint(column) {
			cell.Value = val
			return
		}
	}
	sheet.modifiedCells = append(sheet.modifiedCells, &cell)
}

// DeleteRows deletes rows from the sheet
func (sheet *Sheet) DeleteRows(start, end int) (err error) {
	err = sheet.Spreadsheet.service.DeleteRows(sheet, start, end)
	return
}

// DeleteColumns deletes columns from the sheet
func (sheet *Sheet) DeleteColumns(start, end int) (err error) {
	err = sheet.Spreadsheet.service.DeleteColumns(sheet, start, end)
	return
}

// Synchronize reflects the changes of the sheet.
func (sheet *Sheet) Synchronize() (err error) {
	err = sheet.Spreadsheet.service.SyncSheet(sheet)
	return
}

func newCells(maxRow, maxColumn uint) (rows, columns [][]Cell) {
	rows = make([][]Cell, maxRow+1)
	for i := uint(0); i < maxRow+1; i++ {
		rows[i] = make([]Cell, 0, maxColumn+1)
		for t := uint(0); t < maxColumn+1; t++ {
			rows[i] = append(rows[i], Cell{Row: i, Column: t})
		}
	}
	columns = make([][]Cell, maxColumn+1)
	for i := uint(0); i < maxColumn+1; i++ {
		columns[i] = make([]Cell, 0, maxRow+1)
		for t := uint(0); t < maxRow+1; t++ {
			columns[i] = append(columns[i], Cell{Row: t, Column: i})
		}
	}
	return
}

func appendCells(cells [][]Cell, maxRow, maxColumn uint, cell func(uint, uint) Cell) [][]Cell {
	for i := uint(0); i < maxRow; i++ {
		if len(cells) == 0 || int(i) > len(cells)-1 {
			row := make([]Cell, 0, maxColumn)
			for t := uint(0); t < maxColumn; t++ {
				row = append(row, cell(i, t))
			}
			cells = append(cells, row)
		} else {
			for t := uint(len(cells[i]) - 1); t < maxColumn; t++ {
				cells[i] = append(cells[i], cell(i, t))
			}
		}
	}
	return cells
}
