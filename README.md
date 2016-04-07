spreadsheet
===
[![Build Status](https://travis-ci.org/Iwark/spreadsheet.svg?branch=master)](https://travis-ci.org/Iwark/spreadsheet)
[![Coverage Status](https://coveralls.io/repos/github/Iwark/spreadsheet/badge.svg?branch=master)](https://coveralls.io/github/Iwark/spreadsheet?branch=master)
[![GoReport](https://goreportcard.com/badge/Iwark/spreadsheet)](http://goreportcard.com/report/Iwark/spreadsheet)
[![GoDoc](https://godoc.org/github.com/Iwark/spreadsheet?status.svg)](https://godoc.org/github.com/Iwark/spreadsheet)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
![Project Status](https://img.shields.io/badge/status-beta-yellow.svg)

Package `spreadsheet` provides access to the Google Sheets API for reading and updating spreadsheets.

Any pull-request is welcome.

## Example

```go
package main

import (
	"fmt"
	"io/ioutil"

	"github.com/Iwark/spreadsheet"
	"golang.org/x/net/context"
	"golang.org/x/oauth2/google"
)

func main() {
	data, _ := ioutil.ReadFile("client_secret.json")
	conf, _ := google.JWTConfigFromJSON(data, spreadsheet.SpreadsheetScope)
	client := conf.Client(context.TODO())

	service := &spreadsheet.Service{Client: client}
	sheets, _ := service.Get("1mYiA2T4_QTFUkAXk0BE3u7snN2o5FgSRqxmRrn_Dzh4")
	ws, _ = sheets.Get(0)
	for _, row := range ws.Rows {
		for _, cell := range row {
			fmt.Println(cell.Content)
		}
	}

	// Update cell content
	ws.Rows[0][0].Update("hogehoge")

	// Make sure call Synchronize to reflect the changes
	ws.Synchronize()
}
```

## License

Spreadsheet is released under the MIT License.
