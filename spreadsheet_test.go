package spreadsheet

import (
	"io/ioutil"
	"os"
	"testing"

	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
)

var service *Service
var sheets *Worksheets

func TestMain(m *testing.M) {
	data, _ := ioutil.ReadFile("client_secret.json")
	conf, _ := google.JWTConfigFromJSON(data, SpreadsheetScope)
	client := conf.Client(oauth2.NoContext)
	service, _ = New(client)
	sheets, _ = service.Sheets.Worksheets("1mYiA2T4_QTFUkAXk0BE3u7snN2o5FgSRqxmRrn_Dzh4")
	os.Exit(m.Run())
}

func TestWorksheets(t *testing.T) {
	if sheets.Title != "spreadsheet_example" {
		t.Errorf("Failed to get spreadsheet. got: '%s'", sheets.Title)
	}
}

func TestGet(t *testing.T) {
	ws, _ := sheets.Get(0)
	if ws.Title != "TestSheet" {
		t.Errorf("Failed to get worksheet. got: '%s'", ws.Title)
	}
}

func TestFindById(t *testing.T) {
	_, err := sheets.FindById("od6")
	if err != nil {
		t.Error("Failed to find worksheet. error:", err)
	}
	_, err = sheets.FindById("https://spreadsheets.google.com/feeds/worksheets/1mYiA2T4_QTFUkAXk0BE3u7snN2o5FgSRqxmRrn_Dzh4/private/full/od6")
	if err != nil {
		t.Error("Failed to find worksheet. error:", err)
	}
}

func TestCells(t *testing.T) {
	ws, _ := sheets.Get(0)
	if ws.Rows[0][0].Content != "test" {
		t.Errorf("Failed to get cell. got: '%s'", ws.Rows[0][0].Content)
	}
}

func TestUpdateCell(t *testing.T) {
	ws, _ := sheets.Get(0)
	ws.UpdateCell(ws.Rows[0][1], "Updated")
	ws.Synchronize()
	newSheets, _ := service.Sheets.Worksheets("1mYiA2T4_QTFUkAXk0BE3u7snN2o5FgSRqxmRrn_Dzh4")
	ws, _ = newSheets.Get(0)
	if ws.Rows[0][1].Content != "Updated" {
		t.Error("Failed to update cell")
	}
	ws.UpdateCell(ws.Rows[0][1], "")
	ws.Synchronize()
	newSheets, _ = service.Sheets.Worksheets("1mYiA2T4_QTFUkAXk0BE3u7snN2o5FgSRqxmRrn_Dzh4")
	ws, _ = newSheets.Get(0)
	if ws.Rows[0][1].Content != "" {
		t.Error("Failed to update cell")
	}
}
