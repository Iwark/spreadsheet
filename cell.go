package spreadsheet

import "fmt"

// Cell describes a cell data
type Cell struct {
	Row            uint
	Column         uint
	Value          string
	Note           string
	rawValue       ExtendedValue
	effectiveValue ExtendedValue

	modifiedFields string
}

// Pos returns the cell's position like "A1"
func (cell *Cell) Pos() string {
	return numberToLetter(int(cell.Column)+1) + fmt.Sprintf("%d", cell.Row+1)
}

// RawValue returns the raw value of a cell as entered by a user.
// Cells with formulas, for example, return the formula rather than the value of that formula.
func (cell *Cell) RawValue() ExtendedValue {
	return cell.rawValue
}

// EffectiveValue is the effective value of a cell.
// Cells with formulas will return the value of that formula.
func (cell *Cell) EffectiveValue() ExtendedValue {
	return cell.effectiveValue
}
