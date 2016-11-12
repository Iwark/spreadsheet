package spreadsheet

var letters = [...]string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M",
	"N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

func numberToLetter(num int) (str string) {
	if num == 0 {
		return
	}
	nums := []int{}
	for {
		nums = append(nums, num%26)
		num = num / 26
		if num == 0 {
			break
		}
	}
	for i, j := 0, len(nums)-1; i < j; i, j = i+1, j-1 {
		nums[i], nums[j] = nums[j], nums[i]
	}
	for _, n := range nums {
		str += letters[n-1]
	}
	return
}
