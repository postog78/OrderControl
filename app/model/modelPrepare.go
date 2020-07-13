package model

import (
	"fmt"
	"time"
)

const asciiA = 65

type PreparedDateToAppendIntoSpreadSheet struct {
	SheetName string
	Range     string
	Table     [][]interface{}
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
	return t.Format("02.01.2006")
}

//Prepare возвращает подготовленный массив для записи данных в SpreadSheet гугл из считанных данных отчёта
func Prepare(readerReport []TypeReport) (arrayPrepared []PreparedDateToAppendIntoSpreadSheet) {
	mapSheets := dataPrepare(readerReport)
	count := len(mapSheets)
	arrayPrepared = make([]PreparedDateToAppendIntoSpreadSheet, count)
	i := 0
	for date, val := range mapSheets {
		var prep PreparedDateToAppendIntoSpreadSheet
		prep.Table = val
		prep.SheetName = getSheetNameFromDate(date) + " " + time.Now().Format(time.RFC3339)
		prep.Range = prep.SheetName + "!A1:" + string(asciiA+len(val[0])-1) + fmt.Sprintf("%d", len(val))
		arrayPrepared[i] = prep
		i++
	}
	return arrayPrepared
}
