package spreadsheet

import (
	"io/ioutil"
	"os"
	"testing"

	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
)

var sheets *Worksheets
var ws *Worksheet

func TestMain(m *testing.M) {
	data, _ := ioutil.ReadFile("client_secret.json")
	conf, _ := google.JWTConfigFromJSON(data, SpreadsheetScope)
	client := conf.Client(oauth2.NoContext)
	service, _ := New(client)
	sheets, _ = service.Sheets.Worksheets("1mYiA2T4_QTFUkAXk0BE3u7snN2o5FgSRqxmRrn_Dzh4")
	ws, _ = sheets.Get(0)
	os.Exit(m.Run())
}

func TestWorksheets(t *testing.T) {
	if sheets.Title != "spreadsheet_example" {
		t.Errorf("Failed to get spreadsheet. got: '%s'", sheets.Title)
	}
}

func TestGet(t *testing.T) {
	if ws.Title != "TestSheet" {
		t.Errorf("Failed to get worksheet. got: '%s'", ws.Title)
	}
}

func TestCells(t *testing.T) {
	if ws.Rows[0][0] != "test" {
		t.Errorf("Failed to get cell. got: '%s'", ws.Rows[0][0])
	}
}
