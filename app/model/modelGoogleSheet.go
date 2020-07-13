package model

import (
	"context"
	"errors"
	"fmt"
	"io/ioutil"
	"log"
	"regexp"
	"strconv"
	"strings"
	"time"

	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/sheets/v4"
)

func UpdateGoogleSheetReport(readerReport []TypeReport) {
	// var mapSheetsFromDate map[time.Time][]*sheets.Sheet
	var sheetID string = "1Z7KluzcrGpLu_TVyGUBicwt_mpc1X3FhGlfGluUXh-8"

	// var mapSheet mapGoogleSheetsByMonth
	// var pointMapSheet = &mapSheet

	sheetsService, err := getGoogleSheetSrv()
	if err != nil {
		if err != nil {
			log.Fatalf("Unable to retrieve Sheets client: %v", err)
		}
	}

	// resp := sheetsService.Spreadsheets.Get(sheetID)
	// var ss *sheets.Spreadsheet
	// ss, _ = resp.Do()

	// _ = ss

	ctx := context.Background()
	// How the input data should be interpreted.
	valueInputOption := "USER_ENTERED" // TODO: Update placeholder value.
	// The new values to apply to the spreadsheet.
	//	rangeData := "sheet1!A1:B3"
	//	values := [][]interface{}{{"sample_A1", "sample_B1"}, {"sample_A2", "sample_B2"}, {"sample_A3", "sample_A3"}}

	//data := []*sheets.ValueRange{} // TODO: Update placeholder value.

	rb := &sheets.BatchUpdateValuesRequest{
		ValueInputOption: valueInputOption,
		//	Data:             data,

		// TODO: Add desired fields of the request body.
	}

	preparedData := Prepare(readerReport)
	for _, prep := range preparedData {

		rb.Data = append(rb.Data, &sheets.ValueRange{
			Range:  prep.Range,
			Values: prep.Table,
		})
	}

	// rangeData = "Лист1!A1:B3"
	// rb.Data = append(rb.Data, &sheets.ValueRange{
	// 	Range:  rangeData,
	// 	Values: values,
	// })

	// _, err = sheetsService.Spreadsheets.Values.BatchUpdate(spreadsheetId, rb).Context(ctx).Do()
	// if err != nil {
	//     log.Fatal(err)
	// }
	// fmt.Println("Done.")

	resp, err := sheetsService.Spreadsheets.Values.BatchUpdate(sheetID, rb).Context(ctx).Do()
	if err != nil {
		log.Fatal(err)
	}

	// TODO: Change code below to process the `resp` object:
	fmt.Printf("%#v\n", resp)

	// for _, googleSheet := range ss.Sheets {
	// 	key, err := getTimeFromTitleSheet(googleSheet.Properties.Title)
	// 	if err == nil {
	// 		(&mapSheet).Insert(key, googleSheet)
	// 		// pointMapSheet.Insert(key, googleSheet)
	// 	}
	// }

	// // resp_sheets := resp.Get("sheets")
	// // fmt.Printf("%#v\n", resp_sheets)
	// // if err != nil {
	// // 	log.Fatal(err)
	// // }

	// return mapSheet
}

// //CreateGoogleSheet Создаёт новый гугл документ со считанными данными
// func CreateGoogleSheet(readerReport []TypeReport) {

// 	// spreadsheet := service.spreadsheets().create(body: {
// 	// 	'properties': {'title': 'Сие есть название документа', 'locale': 'ru_RU'},
// 	// 	'sheets': [{'properties': {'sheetType': 'GRID',
// 	// 							   'sheetId': 0,
// 	// 							   'title': 'Сие есть название листа',
// 	// 							   'gridProperties': {'rowCount': 8, 'columnCount': 5}}}]
// 	// }).execute()

// 	// _ = spreadsheet
// 	//Прочитать гугл документ

// 	// Spreadsheet spreadsheet := Spreadsheet()
// 	//     .setProperties(new SpreadsheetProperties()
// 	//             .setTitle(title));

// 	sheetsService, err := getGoogleSheetSrv()
// 	if err != nil {
// 		if err != nil {
// 			log.Fatalf("Unable to retrieve Sheets client: %v", err)
// 		}
// 	}

// 	// _ = srv
// 	// spreadsheet := {
// 	// 	'properties': {
// 	// 		'title': title
// 	// 	}
// 	// }
// 	rb := &sheets.Spreadsheet{
// 		Properties: &sheets.SpreadsheetProperties{
// 			Title:  "Что-то новое",
// 			Locale: "ru_RU",
// 		},
// 	}

// 	ctx := context.Background()
// 	resp, err := sheetsService.Spreadsheets.Create(rb).Context(ctx).Do()
// 	if err != nil {
// 		log.Fatal(err)
// 	}

// 	resp.SpreadsheetId.
// 		// TODO: Change code below to process the `resp` object:
// 		fmt.Printf("%#v\n", resp)
// 	// spreadsheet = srv.spreadsheets().new()create(body=spreadsheet,
// 	//  									fields='spreadsheetId').execute()
// 	// print('Spreadsheet ID: {0}'.format(spreadsheet.get('spreadsheetId')))
// }

// WriteDataToGoogleSheet - функция пишет подоготовленные данные из массива TypeReport в гугл документ
func WriteDataToGoogleSheet(readerReport []TypeReport) {
	//Прежде чем записать данные, надо найти, где они находятся
	//Они находятся на вкладке со специальным именем и месяцем.
	//Название = Месяц + год "Апрель 2020". С любой стороны могут быть данные от имени
	//Каждую дату надо проверять  - на какую вкладку она должна попасть.
	//Получается, что каждая дата не может однозначно лежать только на одной закладке
	//Пока не придуман оптимальный алгоритм, определим функцию, которая по дате будет возвращать список возможных листов
	//эта функция будет также смотреть на месяц назад на всякий случай

	//Прочитать гугл документ
	srv, err := getGoogleSheetSrv()
	if err != nil {
		if err != nil {
			log.Fatalf("Unable to retrieve Sheets client: %v", err)
		}
	}

	//Определить и запомнить список листов
	mapGoolgeSheets := getSheetsList(srv)

	//Инициализируем настройки листов
	mapSheetsSettings := make(map[*sheets.Sheet]sheetSetting)

	dataToSearch := make([]dataToSearchInGoogleSheets, len(readerReport))
	for i, rep := range readerReport {
		sheetsOfDate := mapGoolgeSheets.getSheetsForDateWithMonthBefore(rep.Date)
		dataToSearch[i].rep = rep
		dataToSearch[i].sheets = sheetsOfDate

		for _, sh := range sheetsOfDate {
			addSettings(mapSheetsSettings, sh)
			fmt.Printf("Date %v Title %#v; ", rep.Date, sh.Properties.Title)
		}
		fmt.Println()

	}
	//SearchData(dataToSearch)
	//Запоминаем путь в гугл шит для каждой позиции. //Если позиций несколько, то остаётся только первая найденная.
	//Пишем данные
}

func addSettings(m map[*sheets.Sheet]sheetSetting, sh *sheets.Sheet) {
	_, ok := m[sh]
	if ok {
		return
	}
	// data := sh.Data;
	//fmt.Println("Columns ", sh.getLastRow())
	// sh.ColumnGroups
	// for data.
}

//SearchData ищет в гугл документе данные, которым нужно сделать апдейт
func SearchData(srv *sheets.Service, dataToSearch dataToSearchInGoogleSheets) {
	//Продажа Апрель Заявка E, Масса K, Дата L
	//Продажа Апрель Заявка _, Масса J, Дата K
	//Продажа Джет Заявка E, Масса K, Дата L
	//Продажа ВЛПДС Заявка _, Масса K, Дата L
	//Продажа Март Регионы E, K, L

}

func getMonthBefore(t time.Time) time.Time {
	return time.Date(t.Year(), t.Month()-1, 1, 0, 0, 0, 0, time.UTC)
}

func (m mapGoogleSheetsByMonth) getSheetsForDateWithMonthBefore(date time.Time) []*sheets.Sheet {
	sheetsOfDate := m.getSheetsOfDate(date)
	sheetsOfDateBefore := m.getSheetsOfDate(getMonthBefore(date))
	for _, oneSheet := range sheetsOfDateBefore {
		sheetsOfDate = append(sheetsOfDate, oneSheet)
	}
	return sheetsOfDate
}

func (m mapGoogleSheetsByMonth) getSheetsOfDate(t time.Time) []*sheets.Sheet {
	firstday := time.Date(t.Year(), t.Month(), 1, 0, 0, 0, 0, time.UTC)
	// fmt.Println("Извлечение", firstday)
	sheets, _ := m.mapSheets[firstday]
	return sheets
}

func getGoogleSheetSrv() (*sheets.Service, error) {

	b, err := ioutil.ReadFile("diversityshipments-31d3872ee94b.json")

	if err != nil {
		log.Fatalf("Unable to read client secret file: %v", err)
	}

	// If modifying these scopes, delete your previously saved token.json.
	// config, err := google.ConfigFromJSON(b, "https://www.googleapis.com/auth/spreadsheets.readonly")
	ctx := context.Background()
	// creds, err := google.CredentialsFromJSON(ctx, b, "https://www.googleapis.com/auth/spreadsheets")
	creds, err := google.CredentialsFromJSON(ctx, b, "https://www.googleapis.com/auth/spreadsheets")
	if err != nil {
		log.Fatal(err)
	}

	client := oauth2.NewClient(ctx, creds.TokenSource) //credentials. getClient(config)

	srv, err := sheets.New(client)
	if err != nil {
		log.Fatalf("Unable to retrieve Sheets client: %v", err)
	}
	return srv, nil
}

func getSheetsList(sheetsService *sheets.Service) mapGoogleSheetsByMonth {
	// True if grid data should be returned.
	// This parameter is ignored if a field mask was set in the request.
	// includeGridData := false // TODO: Update placeholder value.
	// ctx := context.Background()

	// The ranges to retrieve from the spreadsheet.
	// ranges := []string{} // TODO: Update placeholder value.

	//resp, err := sheetsService.Spreadsheets.Get(spreadsheetId).Ranges(ranges...).IncludeGridData(includeGridData).Context(ctx).Do()
	resp := sheetsService.Spreadsheets.Get(spreadsheetID)
	var ss *sheets.Spreadsheet
	ss, _ = resp.Do()
	var mapSheet mapGoogleSheetsByMonth
	// var pointMapSheet = &mapSheet

	for _, googleSheet := range ss.Sheets {
		key, err := getTimeFromTitleSheet(googleSheet.Properties.Title)
		if err == nil {
			(&mapSheet).Insert(key, googleSheet)
			// pointMapSheet.Insert(key, googleSheet)
		}
	}
	// // resp_sheets := resp.Get("sheets")
	// // fmt.Printf("%#v\n", resp_sheets)
	// // if err != nil {
	// // 	log.Fatal(err)
	// // }

	return mapSheet
	// fmt.Printf("%#v\n", len(mapSheet.mapSheets))
	// fmt.Printf("%#v\n", mapSheet)

	// Prints the names and majors of students in a sample spreadsheet:
	// https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
	// spreadsheetId := "1_EokmBJWtNKYMbQZEeuTrzk7fwLy5z1H-IbzR1v2q14"
	// spreadsheetId := "183IDyrxg5PczVLewXMronbduZy50ukDiqNUGgnloqQQ"
	// readRange := "МОСГОРТРАНС!A1:E5"
	// resp, err := srv.Spreadsheets.Values.Get(spreadsheetId, readRange).Do()
	// if err != nil {
	// 	log.Fatalf("Unable to retrieve data from sheet: %v", err)
	// }

	// if len(resp.Values) == 0 {
	// 	fmt.Println("No data found.")
	// } else {
	// 	fmt.Println("Name, Major:")
	// 	for _, row := range resp.Values {
	// 		// Print columns A and E, which correspond to indices 0 and 4.
	// 		fmt.Printf("%s\n", row[0])
	// 	}
	// }

	// writeRange := "A1"

	// var vr sheets.ValueRange

	// myval := []interface{}{"One", "Two", "Three"}
	// vr.Values = append(vr.Values, myval)

	// _, err = srv.Spreadsheets.Values.Update(spreadsheetId, writeRange, &vr).ValueInputOption("RAW").Do()
	// if err != nil {
	// 	log.Fatalf("Unable to retrieve data from sheet. %v", err)
	// }
}

// resp, err := srv.Spreadsheets.Values.Get(spreadsheetId, readRange).Do()

// // Prints the names and majors of students in a sample spreadsheet:
// // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
// // spreadsheetId := "1_EokmBJWtNKYMbQZEeuTrzk7fwLy5z1H-IbzR1v2q14"
// spreadsheetId := "183IDyrxg5PczVLewXMronbduZy50ukDiqNUGgnloqQQ"
// readRange := "МОСГОРТРАНС!A1:E5"
// resp, err := srv.Spreadsheets.Values.Get(spreadsheetId, readRange).Do()
// if err != nil {
// 	log.Fatalf("Unable to retrieve data from sheet: %v", err)
// }

// if len(resp.Values) == 0 {
// 	fmt.Println("No data found.")
// } else {
// 	fmt.Println("Name, Major:")
// 	for _, row := range resp.Values {
// 		// Print columns A and E, which correspond to indices 0 and 4.
// 		fmt.Printf("%s\n", row[0])
// 	}
// }

// writeRange := "A1"

// var vr sheets.ValueRange

// myval := []interface{}{"One", "Two", "Three"}
// vr.Values = append(vr.Values, myval)

// _, err = srv.Spreadsheets.Values.Update(spreadsheetId, writeRange, &vr).ValueInputOption("RAW").Do()
// if err != nil {
// 	log.Fatalf("Unable to retrieve data from sheet. %v", err)
// }

func getTimeFromTitleSheet(title string) (retTime time.Time, err error) {
	if len(title) == 0 {
		err = errors.New("No title for sheet")
		return
	}
	re := regexp.MustCompile(`(?i)(Январь|Февраль|Март|Апрель|Май|Июнь|Июль|Август|Сентябрь|Октябрь|Ноябрь|Декабрь) (\d{4})`)
	mapMonth := map[string]time.Month{
		"январь": time.January, "февраль": time.February, "март": time.March,
		"апрель": time.April, "май": time.May, "июнь": time.June,
		"июль": time.July, "август": time.August, "сентябрь": time.September,
		"октябрь": time.October, "ноябрь": time.November, "декабрь": time.December,
	}
	//На всякий случай берём на три года назад и только с целью разработки
	var matches [][]string = re.FindAllStringSubmatch(title, -1)
	if !(len(matches) == 1 && len(matches[0]) == 3) {
		err = errors.New("This is not a our template")
		return
	}
	year, _ := strconv.Atoi(matches[0][2])
	retTime = time.Date(year, mapMonth[strings.ToLower(matches[0][1])], 1, 0, 0, 0, 0, time.UTC)
	return
	// matches[1]
	// Продажа Февраль 2020 (РЕГИОНЫ)
	// ВЛПДС ПРОДАЖА (Январь 2020)
	// Загрузка_марийка
	// Продажа Декабрь 2019 (РЕГИОНЫ)

}
