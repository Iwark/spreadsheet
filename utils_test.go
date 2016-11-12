package spreadsheet

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestNumberToLetter(t *testing.T) {
	assert := assert.New(t)
	assert.Equal("C", numberToLetter(3))
	assert.Equal("AB", numberToLetter(28))
	assert.Equal("AAC", numberToLetter(705))
}
