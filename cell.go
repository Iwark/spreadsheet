package spreadsheet

import "fmt"

// Cell describes a cell data
type Cell struct {
	Row    uint
	Column uint
	Value  string
}

// Pos returns the cell's position like "A1"
func (cell *Cell) Pos() string {
	return numberToLetter(int(cell.Column)+1) + fmt.Sprintf("%d", cell.Row+1)
}
