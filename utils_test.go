package spreadsheet

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestNumberToLetter(t *testing.T) {
	assert := assert.New(t)
	assert.Equal("C", NumberToLetter(3))
	assert.Equal("Z", NumberToLetter(26))
	assert.Equal("AB", NumberToLetter(28))
	assert.Equal("AZ", NumberToLetter(52))
	assert.Equal("AAC", NumberToLetter(705))
	assert.Equal("YZ", NumberToLetter(676))
	assert.Equal("ZA", NumberToLetter(677))
}

func BenchmarkNumberToLetter(b *testing.B) {
	b.ReportAllocs()
	for i := 0; i < b.N; i++ {
		_ = numberToLetter(i)
	}
}
