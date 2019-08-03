package spreadsheet

import (
	"fmt"
	"strings"
)

// Cell describes a cell data
type Cell struct {
	Row    uint
	Column uint
	Value  string
	Note   string

	modifiedFields string
}

// Pos returns the cell's position like "A1"
func (cell *Cell) Pos() string {
	return numberToLetter(int(cell.Column)+1) + fmt.Sprintf("%d", cell.Row+1)
}

func (cell *Cell) addModified(field string) {
	if len(cell.modifiedFields) == 0 {
		cell.modifiedFields = field
	} else if strings.Index(cell.modifiedFields, field) == -1 {
		cell.modifiedFields += "," + field
	}
}
