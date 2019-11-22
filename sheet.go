package spreadsheet

import (
	"encoding/json"
	"strings"
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
					Note:   cellData.Note,
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

func (sheet *Sheet) updateCellField(row, column int, updater func(c *Cell) string) {
	if uint(row)+1 > sheet.newMaxRow {
		sheet.newMaxRow = uint(row) + 1
	}
	if uint(column)+1 > sheet.newMaxColumn {
		sheet.newMaxColumn = uint(column) + 1
	}

	var cell *Cell
	if uint(len(sheet.Rows)) < sheet.newMaxRow+1 ||
		uint(len(sheet.Columns)) < sheet.newMaxColumn+1 {
		sheet.Rows = appendCells(sheet.Rows, sheet.newMaxRow, sheet.newMaxColumn, func(i, t uint) Cell {
			return Cell{Row: i, Column: t}
		})
		sheet.Columns = appendCells(sheet.Columns, sheet.newMaxColumn, sheet.newMaxRow, func(i, t uint) Cell {
			return Cell{Row: t, Column: i}
		})
		cell = &Cell{
			Row:    uint(row),
			Column: uint(column),
		}
	} else {
		cellCopy := sheet.Rows[row][column]
		cell = &cellCopy
	}

	var found bool
	for _, modifiedCell := range sheet.modifiedCells {
		if modifiedCell.Row == uint(row) && modifiedCell.Column == uint(column) {
			cell = modifiedCell
			found = true
			break
		}
	}

	tag := updater(cell)
	if len(cell.modifiedFields) == 0 {
		cell.modifiedFields = tag
	} else if strings.Index(cell.modifiedFields, tag) == -1 {
		cell.modifiedFields += "," + tag
	}

	cellVal := *cell
	cellVal.modifiedFields = ""
	sheet.Rows[row][column] = cellVal
	sheet.Columns[column][row] = cellVal

	if !found {
		sheet.modifiedCells = append(sheet.modifiedCells, cell)
	}
}

// Update updates cell changes
func (sheet *Sheet) Update(row, column int, val string) {
	sheet.updateCellField(row, column, func(c *Cell) string {
		c.Value = val
		return "userEnteredValue"
	})
}

// UpdateNote updates a cell's note
func (sheet *Sheet) UpdateNote(row, column int, note string) {
	sheet.updateCellField(row, column, func(c *Cell) string {
		c.Note = note
		return "note"
	})
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
