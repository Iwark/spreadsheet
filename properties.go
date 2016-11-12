package spreadsheet

// Properties is properties of a spreadsheet.
type Properties struct {
	Title      string `json:"title"`
	Locale     string `json:"locale"`
	AutoRecalc string `json:"autoRecalc"`
	TimeZone   string `json:"timezone"`
	// DefaultFormat *CellFormat `defaultFormat`
}
