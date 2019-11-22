package spreadsheet

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io/ioutil"
	"net/http"
	"net/url"
	"sync"
	"time"

	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
)

const (
	baseURL = "https://sheets.googleapis.com/v4"

	// Scope is the API scope for viewing and managing your Google Spreadsheet data.
	// Useful for generating JWT values.
	Scope = "https://spreadsheets.google.com/feeds"

	// SecretFileName is used to get client.
	SecretFileName = "client_secret.json"
)

// NewService makes a new service with the secret file.
func NewService() (s *Service, err error) {
	data, err := ioutil.ReadFile(SecretFileName)
	if err != nil {
		return
	}

	conf, err := google.JWTConfigFromJSON(data, Scope)
	if err != nil {
		return
	}

	s = NewServiceWithClient(conf.Client(oauth2.NoContext))
	return
}

// NewServiceWithClient makes a new service by the client.
func NewServiceWithClient(client *http.Client) *Service {
	return &Service{
		baseURL:                  baseURL,
		client:                   client,
		m:                        new(sync.RWMutex),
		configForSpreadsheetByID: make(map[string]spreadsheetConfig, 0),
	}
}

// Service represents a Sheets API service instance.
// Service is the main entry point into using this package.
type Service struct {
	baseURL                  string
	client                   *http.Client
	m                        *sync.RWMutex
	configForSpreadsheetByID map[string]spreadsheetConfig
}

// CreateSpreadsheet creates a spreadsheet with the given title
func (s *Service) CreateSpreadsheet(spreadsheet Spreadsheet) (resp Spreadsheet, err error) {
	sheets := make([]map[string]interface{}, 1)
	for s := range spreadsheet.Sheets {
		sheet := spreadsheet.Sheets[s]
		sheets = append(sheets, map[string]interface{}{"properties": map[string]interface{}{"title": sheet.Properties.Title}})
	}
	body, err := s.post("/spreadsheets", map[string]interface{}{
		"properties": map[string]interface{}{
			"title": spreadsheet.Properties.Title,
		},
		"sheets": sheets,
	})
	if err != nil {
		return
	}
	err = json.Unmarshal([]byte(body), &resp)
	if err != nil {
		return
	}
	return s.FetchSpreadsheet(resp.ID)
}

type spreadsheetConfig struct {
	cacheInterval     time.Duration
	lastCachedAt      time.Time
	cachedSpreadsheet Spreadsheet
}

// FetchSpreadsheetOption is the option for FetchSpreadsheet function
type FetchSpreadsheetOption func(*spreadsheetConfig)

// WithCache gives a cacheInterval option for FetchSpreadsheet function
func WithCache(interval time.Duration) FetchSpreadsheetOption {
	return func(config *spreadsheetConfig) {
		config.cacheInterval = interval
	}
}

// FetchSpreadsheet fetches the spreadsheet by the id.
func (s *Service) FetchSpreadsheet(id string, options ...FetchSpreadsheetOption) (spreadsheet Spreadsheet, err error) {
	s.m.RLock()
	config := s.configForSpreadsheetByID[id]
	s.m.RUnlock()
	for _, o := range options {
		o(&config)
	}

	if config.cacheInterval > 0 && time.Now().Sub(config.lastCachedAt.Add(config.cacheInterval)) <= 0 {
		// use cache
		return config.cachedSpreadsheet, nil
	}

	fields := "spreadsheetId,properties.title,sheets(properties,data.rowData.values(formattedValue,note))"
	fields = url.QueryEscape(fields)
	path := fmt.Sprintf("/spreadsheets/%s?fields=%s", id, fields)
	body, err := s.get(path)
	if err != nil {
		return
	}
	err = json.Unmarshal(body, &spreadsheet)
	if err != nil {
		return
	}
	spreadsheet.service = s

	if config.cacheInterval > 0 {
		config.cachedSpreadsheet = spreadsheet
		config.cachedSpreadsheet.cached = true
		config.lastCachedAt = time.Now()
		s.m.Lock()
		s.configForSpreadsheetByID[id] = config
		s.m.Unlock()
	}
	return
}

// ReloadSpreadsheet reloads the spreadsheet
func (s *Service) ReloadSpreadsheet(spreadsheet *Spreadsheet) (err error) {
	newSpreadsheet, err := s.FetchSpreadsheet(spreadsheet.ID)
	if err != nil {
		return
	}
	spreadsheet.Properties = newSpreadsheet.Properties
	spreadsheet.Sheets = newSpreadsheet.Sheets
	return
}

// AddSheet adds a sheet
func (s *Service) AddSheet(spreadsheet *Spreadsheet, sheetProperties SheetProperties) (err error) {
	r, err := newUpdateRequest(spreadsheet)
	if err != nil {
		return
	}
	err = r.AddSheet(sheetProperties).Do()
	if err != nil {
		return
	}
	err = s.ReloadSpreadsheet(spreadsheet)
	return
}

// DuplicateSheet duplicates the contents of a sheet
func (s *Service) DuplicateSheet(spreadsheet *Spreadsheet, sheet *Sheet, index int, title string) (err error) {
	r, err := newUpdateRequest(spreadsheet)
	if err != nil {
		return
	}
	err = r.DuplicateSheet(sheet, index, title).Do()
	if err != nil {
		return
	}
	err = s.ReloadSpreadsheet(spreadsheet)
	return
}

// DeleteSheet deletes the sheet
func (s *Service) DeleteSheet(spreadsheet *Spreadsheet, sheetID uint) (err error) {
	r, err := newUpdateRequest(spreadsheet)
	if err != nil {
		return
	}
	err = r.DeleteSheet(sheetID).Do()
	if err != nil {
		return
	}
	err = s.ReloadSpreadsheet(spreadsheet)
	return
}

// SyncSheet updates sheet
func (s *Service) SyncSheet(sheet *Sheet) (err error) {
	if sheet.newMaxRow > sheet.Properties.GridProperties.RowCount ||
		sheet.newMaxColumn > sheet.Properties.GridProperties.ColumnCount {
		err = s.ExpandSheet(sheet, sheet.newMaxRow, sheet.newMaxColumn)
		if err != nil {
			return
		}
	}
	r, err := newUpdateRequest(sheet.Spreadsheet)
	if err != nil {
		return
	}
	err = r.UpdateCells(sheet).Do()
	if err != nil {
		return
	}
	sheet.modifiedCells = []*Cell{}
	sheet.Properties.GridProperties.RowCount = sheet.newMaxRow
	sheet.Properties.GridProperties.ColumnCount = sheet.newMaxColumn
	return
}

// ExpandSheet expands the range of the sheet
func (s *Service) ExpandSheet(sheet *Sheet, row, column uint) (err error) {
	props := sheet.Properties
	props.GridProperties.RowCount = row
	props.GridProperties.ColumnCount = column

	r, err := newUpdateRequest(sheet.Spreadsheet)
	if err != nil {
		return
	}
	err = r.UpdateSheetProperties(sheet, &props).Do()
	if err != nil {
		return
	}
	sheet.newMaxRow = row
	sheet.newMaxColumn = column
	return
}

// DeleteRows deletes rows from the sheet
func (s *Service) DeleteRows(sheet *Sheet, start, end int) (err error) {
	sheet.Properties.GridProperties.RowCount -= uint(end - start)
	sheet.newMaxRow -= uint(end - start)
	r, err := newUpdateRequest(sheet.Spreadsheet)
	if err != nil {
		return
	}
	err = r.DeleteDimension(sheet, "ROWS", start, end).Do()
	return
}

// DeleteColumns deletes columns from the sheet
func (s *Service) DeleteColumns(sheet *Sheet, start, end int) (err error) {
	sheet.Properties.GridProperties.ColumnCount -= uint(end - start)
	sheet.newMaxRow -= uint(end - start)
	r, err := newUpdateRequest(sheet.Spreadsheet)
	if err != nil {
		return
	}
	err = r.DeleteDimension(sheet, "COLUMNS", start, end).Do()
	return
}

func (s *Service) get(path string) (body []byte, err error) {
	resp, err := s.client.Get(baseURL + path)
	if err != nil {
		return
	}
	body, err = ioutil.ReadAll(resp.Body)
	resp.Body.Close()
	if err != nil {
		return
	}
	err = s.checkError(body)
	return
}

func (s *Service) post(path string, params map[string]interface{}) (body string, err error) {
	reqBody, err := json.Marshal(params)
	if err != nil {
		return
	}
	resp, err := s.client.Post(baseURL+path, "application/json", bytes.NewReader(reqBody))
	if err != nil {
		return
	}
	bytes, err := ioutil.ReadAll(resp.Body)
	resp.Body.Close()
	if err != nil {
		return
	}
	err = s.checkError(bytes)
	if err != nil {
		return
	}
	body = string(bytes)
	return
}

func (s *Service) checkError(body []byte) (err error) {
	var res map[string]interface{}
	err = json.Unmarshal(body, &res)
	if err != nil {
		return
	}
	resErr, hasErr := res["error"].(map[string]interface{})
	if !hasErr {
		return
	}
	code := resErr["code"].(float64)
	message := resErr["message"].(string)
	status := resErr["status"].(string)
	err = fmt.Errorf("error status: %s, code:%d, message: %s", status, int(code), message)
	return
}
