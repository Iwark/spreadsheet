// Package spreadsheet provides access to the Google Spreadsheet.
//
// Usage example:
//
//   import "github.com/Iwark/spreadsheet"
//   ...
//   service, err := spreadsheet.New(oauthHttpClient)
package spreadsheet // import "github.com/Iwark/spreadsheet"

import (
	"encoding/xml"
	"errors"
	"fmt"
	"io/ioutil"
	"net/http"
	"regexp"
	"strings"
	"time"
)

const basePath = "https://spreadsheets.google.com"

const (
	// View and manage your Google Spreadsheet data
	SpreadsheetScope = "https://spreadsheets.google.com/feeds"
)

// New creates a Service object
func New(client *http.Client) (*Service, error) {
	if client == nil {
		return nil, errors.New("client is nil")
	}
	s := &Service{client: client, BasePath: basePath}
	s.Sheets = NewSheetsService(s)
	return s, nil
}

type Service struct {
	client   *http.Client
	BasePath string

	Sheets *SheetsService
}

func (s *Service) fetchAndUnmarshal(url string, v interface{}) error {
	resp, err := s.client.Get(url)
	if err != nil {
		return err
	}
	body, err := ioutil.ReadAll(resp.Body)
	resp.Body.Close()
	if err != nil {
		return err
	}
	err = xml.Unmarshal(body, v)
	if err != nil {
		return err
	}
	return nil
}

func NewSheetsService(s *Service) *SheetsService {
	ss := &SheetsService{s: s}
	return ss
}

type SheetsService struct {
	s *Service
}

// Worksheets returns the Worksheets object of the client
func (ss *SheetsService) Worksheets(key string) (*Worksheets, error) {
	url := fmt.Sprintf("%s/feeds/worksheets/%s/private/full", ss.s.BasePath, key)
	worksheets := &Worksheets{ss: ss}
	err := ss.s.fetchAndUnmarshal(url, &worksheets)
	if err != nil {
		return nil, err
	}
	return worksheets, nil
}

type Worksheets struct {
	ss *SheetsService

	XMLName xml.Name     `xml:"feed"`
	Title   string       `xml:"title"`
	Entries []*Worksheet `xml:"entry"`
}

// Get returns the worksheet of passed index
func (w *Worksheets) Get(i int) (*Worksheet, error) {
	if len(w.Entries) <= i {
		return nil, errors.New(fmt.Sprintf("worksheet of index %d was not found", i))
	}
	ws := w.Entries[i]
	err := ws.Build(w.ss)
	if err != nil {
		return nil, err
	}
	return ws, nil
}

// FindById returns the worksheet of passed id
func (w *Worksheets) FindById(id string) (*Worksheet, error) {
	var validID = regexp.MustCompile(fmt.Sprintf("%s$", id))
	for _, e := range w.Entries {
		if validID.MatchString(e.Id) {
			err := e.Build(w.ss)
			if err != nil {
				return nil, err
			}
			return e, nil
		}
	}
	return nil, errors.New(fmt.Sprintf("worksheet of id %s was not found", id))
}

// FindByTitle returns the worksheet of passed title
func (w *Worksheets) FindByTitle(title string) (*Worksheet, error) {
	for _, e := range w.Entries {
		if e.Title == title {
			err := e.Build(w.ss)
			if err != nil {
				return nil, err
			}
			return e, nil
		}
	}
	return nil, errors.New(fmt.Sprintf("worksheet of title %s was not found", title))
}

type Worksheet struct {
	Id      string    `xml:"id"`
	Updated time.Time `xml:"updated"`
	Title   string    `xml:"title"`
	Content string    `xml:"content"`
	Links   []Link    `xml:"link"`

	ss            *SheetsService
	CellsFeed     string
	MaxRowNum     int
	MaxColNum     int
	Rows          [][]*Cell
	modifiedCells []*Cell
}

func (ws *Worksheet) Build(ss *SheetsService) error {

	ws.ss = ss

	for _, l := range ws.Links {
		if l.Rel == "http://schemas.google.com/spreadsheets/2006#cellsfeed" {
			ws.CellsFeed = l.Href
			break
		}
	}

	var cells *Cells
	err := ws.ss.s.fetchAndUnmarshal(fmt.Sprintf("%s?return-empty=true", ws.CellsFeed), &cells)
	if err != nil {
		return err
	}
	ws.modifiedCells = make([]*Cell, 0)

	for _, cell := range cells.Entries {
		if cell.Pos.Row > ws.MaxRowNum {
			ws.MaxRowNum = cell.Pos.Row
		}
		if cell.Pos.Col > ws.MaxColNum {
			ws.MaxColNum = cell.Pos.Col
		}
	}
	rows := make([][]*Cell, ws.MaxRowNum)
	for i := 0; i < ws.MaxRowNum; i++ {
		rows[i] = make([]*Cell, ws.MaxColNum)
	}
	for _, cell := range cells.Entries {
		rows[cell.Pos.Row-1][cell.Pos.Col-1] = cell
	}
	ws.Rows = rows

	return nil
}

func (ws *Worksheet) UpdateCell(cell *Cell, content string) {
	cell.Content = content
	for _, mc := range ws.modifiedCells {
		if mc.Id == cell.Id {
			return
		}
	}
	ws.modifiedCells = append(ws.modifiedCells, cell)
}

// Synchronize saves the modified cells
func (ws *Worksheet) Synchronize() error {
	feed := `
    <feed xmlns="http://www.w3.org/2005/Atom"
      xmlns:batch="http://schemas.google.com/gdata/batch"
      xmlns:gs="http://schemas.google.com/spreadsheets/2006">
  `
	feed += fmt.Sprintf("<id>%s</id>", ws.CellsFeed)
	for _, mc := range ws.modifiedCells {
		feed += `<entry>`
		feed += fmt.Sprintf("<batch:id>%d, %d</batch:id>", mc.Pos.Row, mc.Pos.Col)
		feed += `<batch:operation type="update"/>`
		feed += fmt.Sprintf("<id>%s</id>", mc.Id)
		feed += fmt.Sprintf("<link rel=\"edit\" type=\"application/atom+xml\" href=\"%s\"/>", mc.EditLink())
		feed += fmt.Sprintf("<gs:cell row=\"%d\" col=\"%d\" inputValue=\"%s\"/>", mc.Pos.Row, mc.Pos.Col, mc.Content)
		feed += `</entry>`
	}
	feed += `</feed>`
	url := fmt.Sprintf("%s/batch", ws.CellsFeed)
	req, err := http.NewRequest("POST", url, strings.NewReader(feed))
	if err != nil {
		return err
	}
	req.Header.Add("Content-Type", "application/atom+xml;charset=utf-8")

	resp, err := ws.ss.s.client.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()
	_, err = ioutil.ReadAll(resp.Body)
	if err != nil {
		return err
	}

	return nil
}

type Link struct {
	Rel  string `xml:"rel,attr"`
	Type string `xml:"type,attr"`
	Href string `xml:"href,attr"`
}

type Cells struct {
	XMLName xml.Name `xml:"feed"`
	Title   string   `xml:"title"`
	Entries []*Cell  `xml:"entry"`
}

type Cell struct {
	Id      string    `xml:"id"`
	Updated time.Time `xml:"updated"`
	Title   string    `xml:"title"`
	Content string    `xml:"content"`
	Links   []Link    `xml:"link"`
	Pos     struct {
		Row int `xml:"row,attr"`
		Col int `xml:"col,attr"`
	} `xml:"cell"`
}

func (c *Cell) EditLink() string {
	for _, l := range c.Links {
		if l.Rel == "edit" {
			return l.Href
		}
	}
	return ""
}
