package spreadsheet

import (
	"encoding/json"
	"errors"
)

// Spreadsheet represents a spreadsheet.
type Spreadsheet struct {
	ID         string     `json:"spreadsheetId"`
	Properties Properties `json:"properties"`
	Sheets     []Sheet    `json:"sheets"`
	// NamedRanges []*NamedRange `json:"namedRanges"`

	service *Service
}

// UnmarshalJSON embeds spreadsheet to sheets.
func (spreadsheet *Spreadsheet) UnmarshalJSON(data []byte) error {
	type Alias Spreadsheet
	a := (*Alias)(spreadsheet)
	if err := json.Unmarshal(data, a); err != nil {
		return err
	}
	for i := range spreadsheet.Sheets {
		spreadsheet.Sheets[i].Spreadsheet = spreadsheet
	}
	return nil
}

// SheetByIndex gets a sheet by the given index.
func (spreadsheet *Spreadsheet) SheetByIndex(index uint) (sheet *Sheet, err error) {
	for i, s := range spreadsheet.Sheets {
		if s.Properties.Index == index {
			sheet = &spreadsheet.Sheets[i]
			return
		}
	}
	err = errors.New("sheet not found by the index")
	return
}

// SheetByID gets a sheet by the given ID.
func (spreadsheet *Spreadsheet) SheetByID(id uint) (sheet *Sheet, err error) {
	for i, s := range spreadsheet.Sheets {
		if s.Properties.ID == id {
			sheet = &spreadsheet.Sheets[i]
			return
		}
	}
	err = errors.New("sheet not found by the id")
	return
}

// SheetByTitle gets a sheet by the given title.
func (spreadsheet *Spreadsheet) SheetByTitle(title string) (sheet *Sheet, err error) {
	for i, s := range spreadsheet.Sheets {
		if s.Properties.Title == title {
			sheet = &spreadsheet.Sheets[i]
			return
		}
	}
	err = errors.New("sheet not found by the title")
	return
}
