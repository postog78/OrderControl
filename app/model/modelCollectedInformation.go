package model

import (
	googlewraper "OrderControl/app/googlewrapper"
	"fmt"
	"log"
	"time"

	"google.golang.org/api/sheets/v4"
)

type CollectedInformation struct {
	gw          *googlewraper.GoogleSheets
	sheetID     string
	excelReport []TypeReport
}

func (ci *CollectedInformation) Init(excelReport []TypeReport) {
	ci.excelReport = excelReport
	ci.sheetID = "1Z7KluzcrGpLu_TVyGUBicwt_mpc1X3FhGlfGluUXh-8"
	ci.gw = new(googlewraper.GoogleSheets)
	ci.gw.Init("diversityshipments-31d3872ee94b.json")
	ci.gw.Start()
}

func (ci *CollectedInformation) Go() {
	arrayPrepared := ci.Prepare()

	// How the input data should be interpreted.
	valueInputOption := "USER_ENTERED" // TODO: Update placeholder value.
	// The new values to apply to the spreadsheet.
	//	rangeData := "sheet1!A1:B3"
	//	values := [][]interface{}{{"sample_A1", "sample_B1"}, {"sample_A2", "sample_B2"}, {"sample_A3", "sample_A3"}}

	//data := []*sheets.ValueRange{} // TODO: Update placeholder value.

	//Добавляем страницы. Считаем, что страницы уникальные и таких в базе ещё нет
	for _, prep := range arrayPrepared {
		ci.gw.InsertSheet(ci.sheetID, prep.SheetName)
	}

	rb := &sheets.BatchUpdateValuesRequest{
		ValueInputOption: valueInputOption,
		//	Data:             data,

		// TODO: Add desired fields of the request body.
	}

	//Добавляем данные на вновь добавленные страницы
	for _, prep := range arrayPrepared {

		rb.Data = append(rb.Data, &sheets.ValueRange{
			Range:  prep.Range,
			Values: prep.Table,
		})
	}

	resp, err := ci.gw.SheetsService.Spreadsheets.Values.BatchUpdate(ci.sheetID, rb).Context(ci.gw.Ctx).Do()
	if err != nil {
		log.Fatal(err)
	}

	fmt.Printf("%#v\n", resp)
}

func dataPrepare(readerReport []TypeReport) map[time.Time][][]interface{} {
	//rangeData := "sheet1!A1:B3"
	// gvalues := [][]interface{}{{"sample_A1", "sample_B1"}, {"sample_A2", "sample_B2"}, {"sample_A3", "sample_A3"}}
	// var values  [][]interface{}
	// {
	// 	{"sample_A1", "sample_B1"},
	// 	{"sample_A2", "sample_B2"},
	// 	{"sample_A3", "sample_A3"}
	// }
	mapSheets := make(map[time.Time][][]interface{})
	for _, rep := range readerReport {
		var row []interface{} = make([]interface{}, 0, 1)
		row = append(row, rep.BasisName, rep.Date, rep.NumOrder, rep.Volume, rep.Weight, rep.Comment, rep.Row)
		values := mapSheets[rep.Date]
		if values == nil {
			values = make([][]interface{}, 0)
		}
		values = append(values, row)
		mapSheets[rep.Date] = values
	}

	return mapSheets
}

func getSheetNameFromDate(t time.Time) string {
	//fmt.Println(fmt.Sprintf("%d", int64(time.Now().Unix())))
	return t.Format("02.01.2006") + " (" + fmt.Sprintf("%d", int64(time.Now().Unix())) + ")"
}

//Prepare возвращает подготовленный массив для записи данных в SpreadSheet гугл из считанных данных отчёта
func (ci *CollectedInformation) Prepare() (arrayPrepared []googlewraper.PreparedDateToAppendIntoSpreadSheet) {
	mapSheets := dataPrepare(ci.excelReport)
	count := len(mapSheets)
	arrayPrepared = make([]googlewraper.PreparedDateToAppendIntoSpreadSheet, count)
	i := 0
	for date, val := range mapSheets {
		var prep googlewraper.PreparedDateToAppendIntoSpreadSheet
		prep.Table = val
		prep.SheetName = getSheetNameFromDate(date)
		prep.Range = prep.SheetName + "!A1:" + string(googlewraper.ASCIISymbolA+len(val[0])-1) + fmt.Sprintf("%d", len(val))
		arrayPrepared[i] = prep
		i++
	}
	return arrayPrepared
}
