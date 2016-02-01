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
	for _, e := range w.Entries {
		if e.Id == id {
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

	ss        *SheetsService
	MaxRowNum int
	MaxColNum int
	Cells     [][]string
}

func (ws *Worksheet) Build(ss *SheetsService) error {
	ws.ss = ss
	xmlCells, err := ws.fetchCells()
	if err != nil {
		return err
	}
	for _, cell := range xmlCells.Entries {
		if cell.Pos.Row > ws.MaxRowNum {
			ws.MaxRowNum = cell.Pos.Row
		}
		if cell.Pos.Col > ws.MaxColNum {
			ws.MaxColNum = cell.Pos.Col
		}
	}
	cells := make([][]string, ws.MaxRowNum)
	for i := 0; i < ws.MaxRowNum; i++ {
		cells[i] = make([]string, ws.MaxColNum)
	}
	for _, cell := range xmlCells.Entries {
		cells[cell.Pos.Row-1][cell.Pos.Col-1] = cell.Content
	}
	ws.Cells = cells

	return nil
}

func (ws *Worksheet) fetchCells() (*Cells, error) {
	var url string
	for _, l := range ws.Links {
		if l.Rel == "http://schemas.google.com/spreadsheets/2006#cellsfeed" {
			url = l.Href
		}
	}
	var cells *Cells
	err := ws.ss.s.fetchAndUnmarshal(url, &cells)
	if err != nil {
		return nil, err
	}
	return cells, nil
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
