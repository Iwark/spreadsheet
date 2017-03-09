package spreadsheet

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestNewCells(t *testing.T) {
	assert := assert.New(t)
	rows, columns := newCells(2, 3)
	assert.Equal(2+1, len(rows))
	assert.Equal(3+1, len(columns))

	assert.Equal(3+1, len(rows[1]))
	assert.Equal(2+1, len(columns[2]))

	assert.Equal(uint(1), rows[1][2].Row)
	assert.Equal(uint(2), rows[1][2].Column)

	assert.Equal(uint(0), columns[2][0].Row)
	assert.Equal(uint(2), columns[2][2].Column)
}
