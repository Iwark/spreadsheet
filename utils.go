package spreadsheet

import (
	"math"
	"strconv"
)

func numberToLetter(num int) string {
	if num <= 0 {
		return ""
	}

	return numberToLetter(int((num-1)/26)) + string(byte(65+(num-1)%26))
}

func cellValueType(val string) string {
	if len(val) == 0 {
		return "stringValue"
	}

	if string(val[0]) == "=" {
		return "formulaValue"
	}
	if _, err := strconv.Atoi(val); err == nil {
		return "numberValue"
	}
	if floatVal, err := strconv.ParseFloat(val, 64); err == nil && isNumericFloat(floatVal) {
		return "numberValue"
	}
	if val == "TRUE" || val == "FALSE" {
		return "boolValue"
	}

	return "stringValue"
}

func isNumericFloat(val float64) bool {
	if math.IsInf(val, 1) || math.IsInf(val, -1) || math.IsNaN(val) {
		return false
	}

	return true
}
