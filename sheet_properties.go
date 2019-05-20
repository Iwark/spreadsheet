package spreadsheet

// SheetProperties is properties of a sheet.
type SheetProperties struct {
	ID             uint           `json:"sheetId,omitempty"`
	Title          string         `json:"title,omitempty"`
	Index          uint           `json:"index,omitempty"`
	SheetType      string         `json:"sheetType,omitempty"`
	GridProperties GridProperties `json:"gridProperties,omitempty"`
	Hidden         bool           `json:"hidden,omitempty"`
	TabColor       TabColor       `json:"tabColor,omitempty"`
	RightToLeft    bool           `json:"rightToLeft,omitempty"`
}
