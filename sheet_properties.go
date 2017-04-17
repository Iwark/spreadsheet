package spreadsheet

// SheetProperties is properties of a sheet.
type SheetProperties struct {
	ID             uint           `json:"sheetId"`
	Title          string         `json:"title"`
	Index          uint           `json:"index"`
	SheetType      string         `json:"sheetType"`
	GridProperties GridProperties `json:"gridProperties"`
	Hidden         bool           `json:"hidden"`
	TabColor       TabColor       `json:"tabColor"`
	RightToLeft    bool           `json:"rightToLeft"`
}
