package spreadsheet

import (
	"errors"
	"regexp"
	"strconv"
)

func NumberToLetter(num int) string {
	if num <= 0 {
		return ""
	}

	return NumberToLetter(int((num-1)/26)) + string(byte(65+(num-1)%26))
}

func LetterToNumber(letter string) (int, error) {
	var total int
	l := len(letter)
	if l >= 14 {
		return 0, errors.New("only references shorter than 14 characters are supported")
	}
	for n, c := range []byte(letter) {
		v := int(byte(c) - 65)
		// On invalid value, just return zero.
		if (v < 0) || (v > 25) {
			return 0, errors.New("expecting upper case A-Z only")
		}
		// Calculate 26 to the power of the place value.
		r := 1
		for i := 1; i < l-n; i++ {
			r *= 26
		}
		// Multiply by the character value starting from A=1 and add to total.
		total += r * ((v % 26) + 1)
	}
	return total, nil
}

func ParseReferenceRange(ref string) (sheet string, x1, y1, x2, y2 int, err error) {
	re := regexp.MustCompile("^'?([^!]+?)'?[!]([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)$")
	xyparts := re.FindAllStringSubmatch(ref, -1)
	if xyparts == nil {
		err = errors.New("can't parse sheet reference")
		return
	}
	sheet = xyparts[0][1]
	x1, err = LetterToNumber(xyparts[0][2])
	if err != nil {
		return
	}
	y1, err = strconv.Atoi(xyparts[0][3])
	if err != nil {
		return
	}
	x2, err = LetterToNumber(xyparts[0][4])
	if err != nil {
		return
	}
	y2, err = strconv.Atoi(xyparts[0][5])
	return
}
