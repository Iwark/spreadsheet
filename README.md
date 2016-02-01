spreadsheet
===
[![GoDoc](https://godoc.org/github.com/Iwark/spreadsheet?status.svg)](https://godoc.org/github.com/Iwark/spreadsheet)

Package ``spreadsheet`` is currently under construction.

Any pull-request is welcome.

## Example

```go
package main

import (
  "io/ioutil"

  "github.com/Iwark/spreadsheet"
  "golang.org/x/oauth2"
  "golang.org/x/oauth2/google"
)

func main(){
  data, _ := ioutil.ReadFile("client_secret.json")
  conf, _ := google.JWTConfigFromJSON(data, spreadsheet.SpreadsheetScope)
  client := conf.Client(oauth2.NoContext)
  service, _ := spreadsheet.New(client)
  sheets, _ := service.Sheets.Worksheets("1mYiA2T4_QTFUkAXk0BE3u7snN2o5FgSRqxmRrn_Dzh4")
  ws, _ = sheets.Get(0)
  for _, row := range ws.Rows {
    for _, cell := range row {
      fmt.Println(cell)
    }
  }
}
```

## License

Spreadsheet is released under the MIT License.