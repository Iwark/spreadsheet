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
	"sync"
	"time"
)

const (
	basePath = "https://spreadsheets.google.com"
	docBase  = "https://docs.google.com/spreadsheets"
)

const (
	// SpreadsheetScope is a scope of View and manage your Google Spreadsheet data
	SpreadsheetScope = "https://spreadsheets.google.com/feeds"
)

// SyncCellsAtOnce is a length of cells to synchronize at once
var SyncCellsAtOnce = 1000

// MaxConnections is the number of max concurrent connections
var MaxConnections = 300

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
	Links   []Link       `xml:"link"`
	Entries []*Worksheet `xml:"entry"`
}

// AddWorksheet adds worksheet
func (ws *Worksheets) AddWorksheet(title string, rowCount, colCount int) error {

	var url string
	for _, l := range ws.Links {
		if l.Rel == "http://schemas.google.com/g/2005#post" {
			url = l.Href
			break
		}
	}
	if url == "" {
		return errors.New("URL not found")
	}

	entry := `<entry xmlns="http://www.w3.org/2005/Atom" xmlns:gs="http://schemas.google.com/spreadsheets/2006">` +
		"<title>" + title + "</title>" +
		fmt.Sprintf("<gs:rowCount>%d</gs:rowCount>", rowCount) +
		fmt.Sprintf("<gs:colCount>%d</gs:colCount>", colCount) +
		`</entry>`

	req, err := http.NewRequest("POST", url, strings.NewReader(entry))
	if err != nil {
		return err
	}
	req.Header.Add("Content-Type", "application/atom+xml;charset=utf-8")

	resp, err := ws.ss.s.client.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()
	body, err := ioutil.ReadAll(resp.Body)
	if err != nil {
		return err
	}

	added := &Worksheet{}
	err = xml.Unmarshal(body, added)
	if err != nil {
		return err
	}
	ws.Entries = append(ws.Entries, added)

	return nil
}

// Get returns the worksheet of passed index
func (w *Worksheets) Get(i int) (*Worksheet, error) {
	if len(w.Entries) <= i {
		return nil, errors.New(fmt.Sprintf("worksheet of index %d was not found", i))
	}
	ws := w.Entries[i]
	err := ws.build(w)
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
			err := e.build(w)
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
			err := e.build(w)
			if err != nil {
				return nil, err
			}
			return e, nil
		}
	}
	return nil, errors.New(fmt.Sprintf("worksheet of title %s was not found", title))
}

// ExistsTitled returns whether there is a sheet titlted given parameter
func (w *Worksheets) ExistsTitled(title string) bool {
	for _, e := range w.Entries {
		if e.Title == title {
			return true
		}
	}
	return false
}

type Worksheet struct {
	Id      string    `xml:"id"`
	Updated time.Time `xml:"updated"`
	Title   string    `xml:"title"`
	Content string    `xml:"content"`
	Links   []Link    `xml:"link"`

	ws            *Worksheets
	CellsFeed     string
	EditLink      string
	CSVLink       string
	MaxRowNum     int
	MaxColNum     int
	Rows          [][]*Cell
	modifiedCells []*Cell
}

// DocsURL is a URL to the google docs spreadsheet (human readable)
func (w *Worksheet) DocsURL() string {
	r := regexp.MustCompile(`/d/(.*?)/export\?gid=(\d+)`)
	group := r.FindSubmatch([]byte(w.CSVLink))
	if len(group) < 3 {
		return ""
	}
	key := string(group[1])
	gid := string(group[2])
	return fmt.Sprintf("%s/d/%s/edit#gid=%s", docBase, key, gid)
}

func (ws *Worksheet) build(w *Worksheets) error {

	ws.ws = w

	for _, l := range ws.Links {
		switch l.Rel {
		case "http://schemas.google.com/spreadsheets/2006#cellsfeed":
			ws.CellsFeed = l.Href
		case "edit":
			ws.EditLink = l.Href
		case "http://schemas.google.com/spreadsheets/2006#exportcsv":
			ws.CSVLink = l.Href
		default:
		}
	}

	var cells *Cells
	err := ws.ws.ss.s.fetchAndUnmarshal(fmt.Sprintf("%s?return-empty=true", ws.CellsFeed), &cells)
	if err != nil {
		return err
	}
	ws.modifiedCells = make([]*Cell, 0)

	for _, cell := range cells.Entries {
		cell.ws = ws
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

func (ws *Worksheet) Destroy() error {
	req, err := http.NewRequest("DELETE", ws.EditLink, nil)
	if err != nil {
		return err
	}

	resp, err := ws.ws.ss.s.client.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()
	_, err = ioutil.ReadAll(resp.Body)
	if err != nil {
		return err
	}

	for i, e := range ws.ws.Entries {
		if e.Id == ws.Id {
			ws.ws.Entries = append(ws.ws.Entries[:i], ws.ws.Entries[i+1:]...)
			break
		}
	}

	return nil
}

// Synchronize saves the modified cells
func (ws *Worksheet) Synchronize() error {

	var wg sync.WaitGroup
	c := make(chan int, MaxConnections)
	mCells := ws.modifiedCells
	target := []*Cell{}
	errors := []error{}
	for len(mCells) > 0 {
		wg.Add(1)
		if len(mCells) >= SyncCellsAtOnce {
			target = mCells[:SyncCellsAtOnce]
			mCells = mCells[SyncCellsAtOnce:]
		} else {
			target = mCells[:len(mCells)]
			mCells = []*Cell{}
		}
		go func(s chan int, cells []*Cell) {
			defer wg.Done()
			s <- 1
			err := ws.synchronize(cells)
			if err != nil {
				errors = append(errors, err)
			}
			<-s
		}(c, target)
	}
	wg.Wait()
	close(c)
	if len(errors) > 0 {
		return errors[0]
	}
	return nil
}

type GSCell struct {
	XMLName    xml.Name `xml:"gs:cell"`
	InputValue string   `xml:"inputValue,attr"`
	Row        int      `xml:"row,attr"`
	Col        int      `xml:"col,attr"`
}

func (ws *Worksheet) synchronize(cells []*Cell) error {
	feed := `
    <feed xmlns="http://www.w3.org/2005/Atom"
      xmlns:batch="http://schemas.google.com/gdata/batch"
      xmlns:gs="http://schemas.google.com/spreadsheets/2006">
  `
	feed += fmt.Sprintf("<id>%s</id>", ws.CellsFeed)
	for _, mc := range cells {
		feed += `<entry>`
		feed += fmt.Sprintf("<batch:id>%d, %d</batch:id>", mc.Pos.Row, mc.Pos.Col)
		feed += `<batch:operation type="update"/>`
		feed += fmt.Sprintf("<id>%s</id>", mc.Id)
		feed += fmt.Sprintf("<link rel=\"edit\" type=\"application/atom+xml\" href=\"%s\"/>", mc.EditLink())
		cell := GSCell{InputValue: mc.Content, Row: mc.Pos.Row, Col: mc.Pos.Col}
		b, err := xml.Marshal(&cell)
		if err != nil {
			return err
		}
		feed += string(b)
		feed += `</entry>`
	}
	feed += `</feed>`
	url := fmt.Sprintf("%s/batch", ws.CellsFeed)
	req, err := http.NewRequest("POST", url, strings.NewReader(feed))
	if err != nil {
		return err
	}
	req.Header.Add("Content-Type", "application/atom+xml;charset=utf-8")

	resp, err := ws.ws.ss.s.client.Do(req)
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
	ws      *Worksheet
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

func (c *Cell) Update(content string) {
	c.Content = content
	for _, mc := range c.ws.modifiedCells {
		if mc.Id == c.Id {
			return
		}
	}
	c.ws.modifiedCells = append(c.ws.modifiedCells, c)
}

func (c *Cell) FastUpdate(content string) {
	c.Content = content
	c.ws.modifiedCells = append(c.ws.modifiedCells, c)
}

func (c *Cell) EditLink() string {
	for _, l := range c.Links {
		if l.Rel == "edit" {
			return l.Href
		}
	}
	return ""
}
