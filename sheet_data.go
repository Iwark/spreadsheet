package spreadsheet

import "encoding/json"

type sheetData struct {
	GridData []GridData `json:"data"`
}

// SheetData is data of the sheet
type SheetData struct {
	sheetData
}

// UnmarshalJSON let SheetData to be unmarshaled
func (d *SheetData) UnmarshalJSON(data []byte) error {
	if err := json.Unmarshal(data, &d.GridData); err != nil {
		return err
	}
	return nil
}
