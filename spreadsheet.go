// Package spreadsheet provides access to the Google Sheets API for reading
// and updating spreadsheets.
//
// Usage example:
//
//   import "github.com/Iwark/spreadsheet"
//   ...
//   service := &spreadsheet.Spreadsheet{Client: oauthHTTPClient}
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
	baseURL = "https://spreadsheets.google.com"
	docBase = "https://docs.google.com/spreadsheets"

	// Scope is the API scope for viewing and managing your Google Spreadsheet data.
	// Useful for generating JWT values.
	Scope = "https://spreadsheets.google.com/feeds"

	// dfltSync is the default number of cells to synchronize at once.
	dfltMaxSync = 1000

	// dfltMaxConns is the default number of max concurrent connections.
	dfltMaxConns = 300
)

// VisibilityState represents a visibility state for a spreadsheet.
type VisibilityState int

const (
	// PrivateVisibility represents a private visibility state for a spreadsheet.  Private
	// spreadsheets require authentication.
	PrivateVisibility VisibilityState = iota

	// PublicVisibility represents a public visibility state for a spreadsheet.  Public
	// spreadsheets can be viewed without authentication.
	PublicVisibility
)

var visibilityName = map[VisibilityState]string{
	PrivateVisibility: "private",
	PublicVisibility:  "public",
}

func (v VisibilityState) String() string {
	return visibilityName[v]
}

// Service represents a Sheets API service instance.  Service is the main entry
// point into using this package.
type Service struct {
	// BaseURL is the base URL used for making API requests.
	// Default is "https://spreadsheets.google.com".
	BaseURL string

	Client *http.Client

	// Maximum number of concurrent connections.
	// Default is 300.
	MaxConns int

	// Maximum number of cells to synchronize at once.
	// Default is 1000.
	MaxSync int

	// private or public.  Default is private.
	Visibility VisibilityState

	// Return all empty cells.
	ReturnEmpty bool
}

// Get returns a spreadsheet with the given ID.
func (s *Service) Get(ID string) (*Spreadsheet, error) {
	if s == nil {
		return nil, errors.New("spreadsheet is nil")
	}

	if s.Client == nil {
		return nil, errors.New("client is nil")
	}

	if s.BaseURL == "" {
		s.BaseURL = baseURL
	}

	if s.MaxSync == 0 {
		s.MaxSync = dfltMaxSync
	}

	if s.MaxConns == 0 {
		s.MaxConns = dfltMaxConns
	}

	url := fmt.Sprintf("%s/feeds/worksheets/%s/%s/full", s.BaseURL, ID, s.Visibility)
	worksheets := &Spreadsheet{s: s}
	err := s.fetchAndUnmarshal(url, &worksheets)
	if err != nil {
		return nil, err
	}
	return worksheets, nil
}

func (s *Service) fetchAndUnmarshal(url string, v interface{}) error {
	resp, err := s.Client.Get(url)
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

// Spreadsheet represents a spreadsheet.  Spreadsheets contain worksheets.
type Spreadsheet struct {
	s *Service

	XMLName    xml.Name     `xml:"feed"`
	Title      string       `xml:"title"`
	Links      []Link       `xml:"link"`
	Worksheets []*Worksheet `xml:"entry"`
}

// NewWorksheet adds a new worksheet.
func (ss *Spreadsheet) NewWorksheet(title string, rowCount, colCount int) (*Worksheet, error) {
	var url string
	for _, l := range ss.Links {
		if l.Rel == "http://schemas.google.com/g/2005#post" {
			url = l.Href
			break
		}
	}
	if url == "" {
		return nil, errors.New("URL not found")
	}

	entry := `<entry xmlns="http://www.w3.org/2005/Atom" xmlns:gs="http://schemas.google.com/spreadsheets/2006">` +
		"<title>" + title + "</title>" +
		fmt.Sprintf("<gs:rowCount>%d</gs:rowCount>", rowCount) +
		fmt.Sprintf("<gs:colCount>%d</gs:colCount>", colCount) +
		`</entry>`

	req, err := http.NewRequest("POST", url, strings.NewReader(entry))
	if err != nil {
		return nil, err
	}
	req.Header.Add("Content-Type", "application/atom+xml;charset=utf-8")

	resp, err := ss.s.Client.Do(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()
	body, err := ioutil.ReadAll(resp.Body)
	if err != nil {
		return nil, err
	}

	added := &Worksheet{}
	err = xml.Unmarshal(body, added)
	if err != nil {
		return nil, err
	}
	ss.Worksheets = append(ss.Worksheets, added)

	return added, nil
}

// Get returns the worksheet at a given index.
func (ss *Spreadsheet) Get(i int) (*Worksheet, error) {
	if i > len(ss.Worksheets)-1 {
		return nil, fmt.Errorf("worksheet of index %d was not found", i)
	}

	ws := ss.Worksheets[i]
	if err := ws.build(ss); err != nil {
		return nil, err
	}
	return ws, nil
}

// FindByID returns the worksheet of passed id.
func (ss *Spreadsheet) FindByID(id string) (*Worksheet, error) {
	s := "/" + id
	for _, e := range ss.Worksheets {
		if strings.HasSuffix(s, e.ID) {
			if err := e.build(ss); err != nil {
				return nil, err
			}
			return e, nil
		}
	}
	return nil, fmt.Errorf("worksheet of id %q was not found", id)
}

// FindByTitle returns the worksheet of passed title.
func (ss *Spreadsheet) FindByTitle(title string) (*Worksheet, error) {
	for _, e := range ss.Worksheets {
		if e.Title == title {
			err := e.build(ss)
			if err != nil {
				return nil, err
			}
			return e, nil
		}
	}
	return nil, fmt.Errorf("worksheet of title %s was not found", title)
}

// ExistsTitled returns whether there is a sheet titlted given parameter
func (ss *Spreadsheet) ExistsTitled(title string) bool {
	for _, e := range ss.Worksheets {
		if e.Title == title {
			return true
		}
	}
	return false
}

// A Worksheet represents a worksheet inside a spreadsheet.
type Worksheet struct {
	ID      string    `xml:"id"`
	Updated time.Time `xml:"updated"`
	Title   string    `xml:"title"`
	Content string    `xml:"content"`
	Links   []Link    `xml:"link"`

	ss            *Spreadsheet
	CellsFeed     string
	EditLink      string
	CSVLink       string
	MaxRowNum     int
	MaxColNum     int
	Rows          [][]*Cell
	modifiedCells []*Cell
}

// DocsURL is a URL to the google docs spreadsheet (human readable)
func (ws *Worksheet) DocsURL() string {
	r := regexp.MustCompile(`/d/(.*?)/export\?gid=(\d+)`)
	group := r.FindSubmatch([]byte(ws.CSVLink))
	if len(group) < 3 {
		return ""
	}
	key := string(group[1])
	gid := string(group[2])
	return fmt.Sprintf("%s/d/%s/edit#gid=%s", docBase, key, gid)
}

func (ws *Worksheet) setLinks() {
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
}

func (ws *Worksheet) build(ss *Spreadsheet) error {
	ws.ss = ss
	ws.setLinks()

	var cells *Cells
	err := ws.ss.s.fetchAndUnmarshal(fmt.Sprintf("%s?return-empty=%v", ws.CellsFeed, ws.ss.s.ReturnEmpty), &cells)
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

// Destroy deletes a worksheet.
func (ws *Worksheet) Destroy() error {
	req, err := http.NewRequest("DELETE", ws.EditLink, nil)
	if err != nil {
		return err
	}

	resp, err := ws.ss.s.Client.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()
	_, err = ioutil.ReadAll(resp.Body)
	if err != nil {
		return err
	}

	for i, e := range ws.ss.Worksheets {
		if e.ID == ws.ID {
			ws.ss.Worksheets = append(ws.ss.Worksheets[:i], ws.ss.Worksheets[i+1:]...)
			break
		}
	}

	return nil
}

// Synchronize saves the modified cells.
func (ws *Worksheet) Synchronize() error {

	var wg sync.WaitGroup
	c := make(chan int, ws.ss.s.MaxConns)
	mCells := ws.modifiedCells
	target := []*Cell{}
	errors := []error{}
	for len(mCells) > 0 {
		wg.Add(1)
		if len(mCells) >= ws.ss.s.MaxSync {
			target = mCells[:ws.ss.s.MaxSync]
			mCells = mCells[ws.ss.s.MaxSync:]
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

type gsCell struct {
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
		feed += fmt.Sprintf("<id>%s</id>", mc.ID)
		feed += fmt.Sprintf("<link rel=\"edit\" type=\"application/atom+xml\" href=\"%s\"/>", mc.EditLink())
		cell := gsCell{InputValue: mc.Content, Row: mc.Pos.Row, Col: mc.Pos.Col}
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

	resp, err := ws.ss.s.Client.Do(req)
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

// Link represents a URL link element within the Sheets API.
type Link struct {
	Rel  string `xml:"rel,attr"`
	Type string `xml:"type,attr"`
	Href string `xml:"href,attr"`
}

// Cells represents a group of cells.
type Cells struct {
	XMLName xml.Name `xml:"feed"`
	Title   string   `xml:"title"`
	Entries []*Cell  `xml:"entry"`
}

// A Cell represents an individual cell in a worksheet.
type Cell struct {
	ws      *Worksheet
	ID      string    `xml:"id"`
	Updated time.Time `xml:"updated"`
	Title   string    `xml:"title"`
	Content string    `xml:"content"`
	Links   []Link    `xml:"link"`
	Pos     struct {
		Row int `xml:"row,attr"`
		Col int `xml:"col,attr"`
	} `xml:"cell"`
}

// Update will update the content of the cell.
func (c *Cell) Update(content string) {
	c.Content = content
	for _, mc := range c.ws.modifiedCells {
		if mc.ID == c.ID {
			return
		}
	}
	c.ws.modifiedCells = append(c.ws.modifiedCells, c)
}

// FastUpdate updates the content of the cell and appends the cell to the list
// of modified cells.
func (c *Cell) FastUpdate(content string) {
	c.Content = content
	c.ws.modifiedCells = append(c.ws.modifiedCells, c)
}

// EditLink returns the edit link for the cell.
func (c *Cell) EditLink() string {
	for _, l := range c.Links {
		if l.Rel == "edit" {
			return l.Href
		}
	}
	return ""
}
