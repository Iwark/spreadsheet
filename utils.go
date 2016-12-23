package spreadsheet

func numberToLetter(num int) string {
	if num <= 0 {
		return ""
	}

	return numberToLetter(int((num-1)/26)) + string(byte(65+(num-1)%26))
}
