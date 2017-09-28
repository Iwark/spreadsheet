spreadsheet
===
[![Build Status](https://travis-ci.org/Iwark/spreadsheet.svg?branch=v2)](https://travis-ci.org/Iwark/spreadsheet)
[![Coverage Status](https://coveralls.io/repos/github/Iwark/spreadsheet/badge.svg?branch=v2)](https://coveralls.io/github/Iwark/spreadsheet?branch=v2)
[![GoReport](https://goreportcard.com/badge/Iwark/spreadsheet)](http://goreportcard.com/report/Iwark/spreadsheet)
[![GoDoc](https://godoc.org/gopkg.in/Iwark/spreadsheet.v2?status.svg)](https://godoc.org/gopkg.in/Iwark/spreadsheet.v2)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

Package `spreadsheet` provides fast and easy-to-use access to the Google Sheets API for reading and updating spreadsheets.

Any pull-request is welcome.

## Installation

```
go get gopkg.in/Iwark/spreadsheet.v2
```

## Preparation

This package uses oauth2 client for authentication. You need to get service account key from [Google Developer Console](https://console.developers.google.com/project). Place the ``client_secret.json`` to the root of your project.

## Usage

First you need **service** to start using this package.

```go
data, err := ioutil.ReadFile("client_secret.json")
checkError(err)

conf, err := google.JWTConfigFromJSON(data, spreadsheet.Scope)
checkError(err)

client := conf.Client(context.TODO())
service := spreadsheet.NewServiceWithClient(client)
```

Or there is a shortcut which does the same things:

```go
service, err := spreadsheet.NewService()
```

### Fetching a spreadsheet

```go
spreadsheetID := "1mYiA2T4_QTFUkAXk0BE3u7snN2o5FgSRqxmRrn_Dzh4"
spreadsheet, err := service.FetchSpreadsheet(spreadsheetID)
```

### Create a spreadsheet

```go
ss, err := service.CreateSpreadsheet(spreadsheet.Spreadsheet{
	Properties: spreadsheet.Properties{
		Title: "spreadsheet title",
	},
})
```

### Find a sheet

```go
// get a sheet by the index.
sheet, err := spreadsheet.SheetByIndex(0)

// get a sheet by the ID.
sheet, err := spreadsheet.SheetByID(0)

// get a sheet by the title.
sheet, err := spreadsheet.SheetByTitle("SheetTitle")
```

### Get cells

```go
// get the B1 cell content
sheet.Rows[0][1].Value

// get the A2 cell content
sheet.Columns[0][1].Value
```

### Update cell content

```go
row := 1
column := 2
sheet.Update(row, column, "hogehoge")
sheet.Update(3, 2, "fugafuga")

// Make sure call Synchronize to reflect the changes.
err := sheet.Synchronize()
```

### Expand a sheet

```go
err := service.ExpandSheet(sheet, 20, 10) // Expand the sheet to 20 rows and 10 columns
```

### Delete Rows / Columns

```go
err := sheet.DeleteRows(0, 3) // Delete first three rows in the sheet

err := sheet.DeleteColumns(1, 4) // Delete columns B:D
```

More usage can be found at the [godoc](https://godoc.org/gopkg.in/Iwark/spreadsheet.v2).

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
