spreadsheet
===
[![Build Status](https://travis-ci.org/Iwark/spreadsheet.svg?branch=master)](https://travis-ci.org/Iwark/spreadsheet)
[![Coverage Status](https://coveralls.io/repos/github/Iwark/spreadsheet/badge.svg?branch=master)](https://coveralls.io/github/Iwark/spreadsheet?branch=master)
[![GoReport](https://goreportcard.com/badge/Iwark/spreadsheet)](http://goreportcard.com/report/Iwark/spreadsheet)
[![GoDoc](https://godoc.org/github.com/Iwark/spreadsheet?status.svg)](https://godoc.org/github.com/Iwark/spreadsheet)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
![Project Status](https://img.shields.io/badge/status-beta-yellow.svg)

Package `spreadsheet` provides fast and easy-to-use access to the Google Sheets API for reading and updating spreadsheets.

Any pull-request is welcome.

## Installation

```
go get gopkg.in/Iwark/spreadsheet.v2
```

## Example

```go
package main

import (
	"fmt"
	"io/ioutil"

	"gopkg.in/Iwark/spreadsheet.v2"
	"golang.org/x/net/context"
	"golang.org/x/oauth2/google"
)

func main() {
	data, err := ioutil.ReadFile("client_secret.json")
	checkError(err)
	conf, err := google.JWTConfigFromJSON(data, spreadsheet.Scope)
	checkError(err)
	client := conf.Client(context.TODO())

	service := spreadsheet.NewServiceWithClient(client)
	spreadsheet, err := service.FetchSpreadsheet("1mYiA2T4_QTFUkAXk0BE3u7snN2o5FgSRqxmRrn_Dzh4")
	checkError(err)
	sheet, err := spreadsheet.SheetByIndex(0)
	checkError(err)
	for _, row := range sheet.Rows {
		for _, cell := range row {
			fmt.Println(cell.Value)
		}
	}

	// Update cell content
	sheet.Update(0, 0, "hogehoge")

	// Make sure call Synchronize to reflect the changes
	err = sheet.Synchronize()
	checkError(err)
}

func checkError(err error) {
	if err != nil {
		panic(err.Error())
	}
}
```

## License

Spreadsheet is released under the MIT License.
