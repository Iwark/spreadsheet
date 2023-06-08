package spreadsheet

import (
	"math"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestNumberToLetter(t *testing.T) {
	assert := assert.New(t)
	assert.Equal("C", numberToLetter(3))
	assert.Equal("Z", numberToLetter(26))
	assert.Equal("AB", numberToLetter(28))
	assert.Equal("AZ", numberToLetter(52))
	assert.Equal("AAC", numberToLetter(705))
	assert.Equal("YZ", numberToLetter(676))
	assert.Equal("ZA", numberToLetter(677))
}

func TestCellValueType(t *testing.T) {
	assert := assert.New(t)
	assert.Equal("stringValue", cellValueType(""))
	assert.Equal("formulaValue", cellValueType("=ABS(-2)"))
	assert.Equal("numberValue", cellValueType("-2"))
	assert.Equal("numberValue", cellValueType("-2.23333"))
	assert.Equal("boolValue", cellValueType("TRUE"))
	assert.Equal("stringValue", cellValueType("test"))
	assert.Equal("stringValue", cellValueType("inf"))
	assert.Equal("stringValue", cellValueType("Infinity"))
	assert.Equal("stringValue", cellValueType("-inf"))
	assert.Equal("stringValue", cellValueType("-Infinity"))
	assert.Equal("stringValue", cellValueType("NaN"))
}

func TestIsNumericFloat(t *testing.T) {
	assert := assert.New(t)
	assert.True(isNumericFloat(-2.23333))
	assert.True(isNumericFloat(0.01234))
	assert.False(isNumericFloat(math.Inf(1)))
	assert.False(isNumericFloat(math.Inf(-1)))
	assert.False(isNumericFloat(math.NaN()))
}

func BenchmarkNumberToLetter(b *testing.B) {
	b.ReportAllocs()
	for i := 0; i < b.N; i++ {
		_ = numberToLetter(i)
	}
}
