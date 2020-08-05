package model

import (
	googlewraper "OrderControl/app/googlewrapper"
	"fmt"
	"log"
	"time"

	"google.golang.org/api/sheets/v4"
)

//CollectedInformation структура хранит даннные о подключении и прочитанном списке отгрузок
type CollectedInformation struct {
	gw            *googlewraper.GoogleSheets
	spreadSheetID string
	excelReport   []TypeReport
}

//Init инициализация данных, соединение с сервером Гугла
func (ci *CollectedInformation) Init(excelReport []TypeReport) {
	ci.excelReport = excelReport
	ci.spreadSheetID = "1Z7KluzcrGpLu_TVyGUBicwt_mpc1X3FhGlfGluUXh-8"
	ci.gw = new(googlewraper.GoogleSheets)
	ci.gw.Init("diversityshipments-31d3872ee94b.json")
	ci.gw.Start()
}

func (ci *CollectedInformation) Go() (newSheetID int64) {
	// arrayPrepared := ci.Prepare()
	arrayPrepared := ci.PrepareForOneSheet()

	// How the input data should be interpreted.
	valueInputOption := "USER_ENTERED" // TODO: Update placeholder value.
	// The new values to apply to the spreadsheet.
	//	rangeData := "sheet1!A1:B3"
	//	values := [][]interface{}{{"sample_A1", "sample_B1"}, {"sample_A2", "sample_B2"}, {"sample_A3", "sample_A3"}}

	//data := []*sheets.ValueRange{} // TODO: Update placeholder value.

	//Добавляем страницы. Считаем, что страницы уникальные и таких в базе ещё нет
	mapNewSheetID := make(map[string]int64)
	for _, prep := range arrayPrepared {
		newSheetID = genegateInt64()
		mapNewSheetID[prep.SheetName] = newSheetID
		ci.gw.InsertSheet(ci.spreadSheetID, prep.SheetName, newSheetID)
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

	_, err := ci.gw.SheetsService.Spreadsheets.Values.BatchUpdate(ci.spreadSheetID, rb).Context(ci.gw.Ctx).Do()
	if err != nil {
		log.Fatal(err)
	}

	//resp.
	//Сделаем фильтр
	gridRange := new(sheets.GridRange)

	// gridRange.SheetId = 0
	gridRange.StartRowIndex = 0
	gridRange.EndRowIndex = 5000
	gridRange.StartColumnIndex = 0
	gridRange.EndColumnIndex = 20

	for _, sheetID := range mapNewSheetID {
		gridRange.SheetId = sheetID
		ci.gw.MakeFilter(ci.spreadSheetID, "Фильтр на всё", gridRange)
	}

	return newSheetID
}

func dataPrepareForManySheet(readerReport []TypeReport) map[time.Time][][]interface{} {
	//rangeData := "sheet1!A1:B3"
	// gvalues := [][]interface{}{{"sample_A1", "sample_B1"}, {"sample_A2", "sample_B2"}, {"sample_A3", "sample_A3"}}
	// var values  [][]interface{}
	// {
	// 	{"sample_A1", "sample_B1"},
	// 	{"sample_A2", "sample_B2"},
	// 	{"sample_A3", "sample_A3"}
	// }
	mapSheets := make(map[time.Time][][]interface{})
	//Title
	var title []interface{} = make([]interface{}, 0, 1)
	title = append(title, "Базис", "Дата заявки", "Номер заявки", "Вид топлива", "Объем, Литры", "Вес, кг", "Комментарий", "Файл", "Лист", "Номер строки")

	//Устанавливаем заголовки на все страницы. Если страницы нет, значит нет и заголовка
	for _, rep := range readerReport {
		values := mapSheets[rep.Date]
		if values == nil {
			values = make([][]interface{}, 0)
			values = append(values, title)
			mapSheets[rep.Date] = values
		}
	}

	//Теперь заполняем данными
	for _, rep := range readerReport {
		values := mapSheets[rep.Date]
		// if values == nil {
		// 	values = make([][]interface{}, 0)
		// }

		var row []interface{} = make([]interface{}, 0, 1)
		row = append(row, rep.BasisName, rep.Date.Format("02.01.2006"), rep.NumOrder, rep.TypeOfProduct, rep.Volume, rep.Weight, rep.Comment, rep.FileName, rep.SheetName, rep.Row)
		values = append(values, row)
		mapSheets[rep.Date] = values
	}

	return mapSheets
}

func dataPrepareForOneSheet(readerReport []TypeReport) (values [][]interface{}, dateBegin, dateEnd time.Time) {
	//Title
	var title []interface{} = make([]interface{}, 0, 1)
	title = append(title,
		"Базис", "Дата заявки", "Номер заявки", "Вид топлива", "Объем, Литры", "Вес, кг",
		"Комментарий", "Водитель", "Файл", "Лист", "Номер строки")
	values = make([][]interface{}, 0)
	values = append(values, title)

	//Теперь заполняем данными
	//var dateBegin, dateEnd time.Time
	for _, rep := range readerReport {

		if dateEnd.Before(rep.Date) {
			dateEnd = rep.Date
		}

		if dateBegin.IsZero() || dateBegin.After(rep.Date) {
			dateBegin = rep.Date
		}

		var row []interface{} = make([]interface{}, 0, 1)

		row = append(row,
			rep.BasisName, rep.Date.Format("02.01.2006"), rep.NumOrder, rep.TypeOfProduct, rep.Volume,
			rep.Weight, rep.Comment, rep.Driver, rep.FileName, rep.SheetName, rep.Row)
		values = append(values, row)
	}

	// type inside struct {
	// 	values    [][]interface{}
	// 	dateBegin time.Time
	// 	dateEnd   time.Time
	// }

	return values, dateBegin, dateEnd
}

func getSheetNameFromDate(t time.Time) string {
	//fmt.Println(fmt.Sprintf("%d", int64(time.Now().Unix())))
	return t.Format("02.01.2006") + " (" + fmt.Sprintf("%d", int64(time.Now().Unix())) + ")"
}

func getSheetNameFromDateBeginEnd(t1, t2 time.Time) string {
	return t1.Format("02.01.2006") + " - " + t2.Format("02.01.2006") + " (" + fmt.Sprintf("%d", int64(time.Now().Unix())) + ")"
}

//Prepare возвращает подготовленный массив для записи данных в SpreadSheet гугл из считанных данных отчёта
func (ci *CollectedInformation) PrepareForManySheets() (arrayPrepared []googlewraper.PreparedDateToAppendIntoSpreadSheet) {
	mapSheets := dataPrepareForManySheet(ci.excelReport)
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

//Prepare возвращает подготовленный массив для записи данных в SpreadSheet гугл из считанных данных отчёта
func (ci *CollectedInformation) PrepareForOneSheet() (arrayPrepared []googlewraper.PreparedDateToAppendIntoSpreadSheet) {
	values, dateBegin, dateEnd := dataPrepareForOneSheet(ci.excelReport)
	arrayPrepared = make([]googlewraper.PreparedDateToAppendIntoSpreadSheet, 1, 1)
	var prep googlewraper.PreparedDateToAppendIntoSpreadSheet
	prep.Table = values
	prep.SheetName = getSheetNameFromDateBeginEnd(dateBegin, dateEnd)
	prep.Range = prep.SheetName + "!A1:" + string(googlewraper.ASCIISymbolA+len(values[0])-1) + fmt.Sprintf("%d", len(values))
	arrayPrepared[0] = prep

	return arrayPrepared
}

func genegateInt64() int64 {
	return int64(time.Now().Unix())
}

func (ci *CollectedInformation) DeleteAllSheets(sheetIDExeptDeletion int64) {

	resp := ci.gw.SheetsService.Spreadsheets.Get(ci.spreadSheetID)
	var ss *sheets.Spreadsheet
	ss, _ = resp.Do()

	for _, googleSheet := range ss.Sheets {
		if googleSheet.Properties.SheetId != sheetIDExeptDeletion {
			ci.gw.DeleteSheet(ci.spreadSheetID, googleSheet.Properties.SheetId)
		}
	}

}
