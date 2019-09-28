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

func benchmarkUpdate(t int, b *testing.B) {
	for f := 0; f < b.N; f++ {
		s := Sheet{}
		b.ReportAllocs()
		for i := 0; i < t; i++ {
			s.Update(i, i, "")
		}
	}
}

func BenchmarkUpdate1(b *testing.B)    { benchmarkUpdate(1, b) }
func BenchmarkUpdate10(b *testing.B)   { benchmarkUpdate(10, b) }
func BenchmarkUpdate100(b *testing.B)  { benchmarkUpdate(100, b) }
func BenchmarkUpdate1000(b *testing.B) { benchmarkUpdate(1000, b) }
