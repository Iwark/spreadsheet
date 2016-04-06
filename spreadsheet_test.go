package spreadsheet

import (
	"io/ioutil"
	"testing"

	"github.com/stretchr/testify/suite"

	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
)

const key = "1mYiA2T4_QTFUkAXk0BE3u7snN2o5FgSRqxmRrn_Dzh4"

type SpreadsheetTestSuite struct {
	suite.Suite
	service *Service
	sheets  *Spreadsheet
}

func (suite *SpreadsheetTestSuite) SetupTest() {
	data, _ := ioutil.ReadFile("client_secret.json")
	conf, _ := google.JWTConfigFromJSON(data, Scope)
	client := conf.Client(oauth2.NoContext)
	suite.service = &Service{Client: client}
	suite.sheets, _ = suite.service.Get(key)
}

func (suite *SpreadsheetTestSuite) TestWorksheets() {
	suite.Equal("spreadsheet_example", suite.sheets.Title)
}

func (suite *SpreadsheetTestSuite) TestGet() {
	ws, _ := suite.sheets.Get(0)
	suite.Equal("TestSheet", ws.Title)
}

func (suite *SpreadsheetTestSuite) TestFindById() {
	_, err := suite.sheets.FindByID("od6")
	suite.Nil(err)
	_, err = suite.sheets.FindByID("https://spreadsheets.google.com/feeds/worksheets/1mYiA2T4_QTFUkAXk0BE3u7snN2o5FgSRqxmRrn_Dzh4/private/full/od6")
	suite.Nil(err)
}

func (suite *SpreadsheetTestSuite) TestFindByTitle() {
	_, err := suite.sheets.FindByTitle("TestSheet")
	suite.Nil(err)
}

func (suite *SpreadsheetTestSuite) TestCells() {
	ws, _ := suite.sheets.Get(0)
	suite.Equal("test", ws.Rows[0][0].Content)
}

func (suite *SpreadsheetTestSuite) TestNewAndDestroyWorksheet() {
	title := "test_adding_sheet"
	_, err := suite.sheets.NewWorksheet(title, 5, 3)
	suite.Nil(err)

	suite.sheets, _ = suite.service.Get(key)
	ws, err := suite.sheets.FindByTitle(title)
	suite.Nil(err)
	suite.Equal(title, ws.Title)

	err = ws.Destroy()
	suite.Nil(err)
	suite.False(suite.sheets.ExistsTitled(title))
}

func (suite *SpreadsheetTestSuite) TestDocsURL() {
	ws, err := suite.sheets.FindByID("od6")
	suite.Nil(err)

	expectedURL := "https://docs.google.com/spreadsheets/d/1mYiA2T4_QTFUkAXk0BE3u7snN2o5FgSRqxmRrn_Dzh4/edit#gid=0"
	url := ws.DocsURL()
	suite.Equal(expectedURL, url)
}

func (suite *SpreadsheetTestSuite) TestUpdateCell() {
	suite.service.ReturnEmpty = true
	defer func() {
		suite.service.ReturnEmpty = false
	}()
	ws, err := suite.sheets.FindByID("od6")
	suite.Nil(err)
	ws.Rows[0][1].Update("Updated")
	ws.Synchronize()

	suite.sheets, _ = suite.service.Get(key)
	ws, _ = suite.sheets.Get(0)
	suite.Equal("Updated", ws.Rows[0][1].Content)
	ws.Rows[0][1].Update("")
	ws.Synchronize()

	suite.sheets, _ = suite.service.Get(key)
	ws, _ = suite.sheets.Get(0)
	suite.Equal("", ws.Rows[0][1].Content)
}

func TestSpreadsheetTestSuite(t *testing.T) {
	suite.Run(t, new(SpreadsheetTestSuite))
}
