package googlewrapper

import (
	"context"
	"fmt"
	"io/ioutil"
	"log"
	"time"

	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/sheets/v4"
)

//GoogleSheets структура описывающая работу в целом с гугл таблицами и хранящая инициализированный сервер
type GoogleSheets struct {
	SheetsService *sheets.Service
	// SpreadsheetID    string
	clientSecretFile string
	Ctx              context.Context
}

//ASCIISymbolA код символа А, чтобы определять для экселя
const ASCIISymbolA = 65

//PreparedDateToAppendIntoSpreadSheet Подготовленные данные для записи в гугл таблицы
type PreparedDateToAppendIntoSpreadSheet struct {
	SheetName string
	Range     string
	Table     [][]interface{}
}

type mapGoogleSheetsByMonth struct {
	mapSheets map[time.Time][]*sheets.Sheet
}

//Init инициализация внутренних переменных
func (g *GoogleSheets) Init(clientSecretFile string) {
	//diversityshipments-31d3872ee94b.json
	g.clientSecretFile = clientSecretFile
}

//Start Запуск соединения с Гуглом
func (g *GoogleSheets) Start() {
	g.startGoogleSheetSrv()
	// sheetsService, err := getGoogleSheetSrv()
	// if err != nil {
	// 	if err != nil {
	// 		log.Fatalf("Unable to retrieve Sheets client: %v", err)
	// 	}
	// }
	// g.sheetsService = sheetsService
	// ctx = context.Background()
}

func (g *GoogleSheets) getSheetsList(spreadsheetID string, getTime func(string) (time.Time, error)) (lst []string) {

	resp := g.SheetsService.Spreadsheets.Get(spreadsheetID)
	var ss *sheets.Spreadsheet
	ss, _ = resp.Do()

	for _, googleSheet := range ss.Sheets {
		lst = append(lst, googleSheet.Properties.Title)
	}
	return
}

//InsertSheet Вставляет страницу в гугл таблицу
func (g *GoogleSheets) InsertSheet(spreadSheetID, name string, sheetID int64) {

	// The spreadsheet to apply the updates to.
	//spreadsheetId := "my-spreadsheet-id" // TODO: Update placeholder value.

	// A list of updates to apply to the spreadsheet.
	// Requests will be applied in the order they are specified.
	// If any request is not valid, no requests will be applied.
	//requests := []*sheets.Request{} // TODO: Update placeholder value.

	req := sheets.Request{
		AddSheet: &sheets.AddSheetRequest{
			Properties: &sheets.SheetProperties{
				Title:   name,
				Index:   0,
				SheetId: sheetID,
			},
		},
	}

	rb := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{&req},
	}

	_, err := g.SheetsService.Spreadsheets.BatchUpdate(spreadSheetID, rb).Context(g.Ctx).Do()
	if err != nil {
		log.Fatal(err)
	}

	// fmt.Printf("%#v\n", resp)

	// {
	// 	requests": [
	// 	  {
	// 		"adSheet": {
	// 		  "properties: {
	// 			"title": "Deposis",
	// 			"gridProperties": {
	// 			  "rowCount": 20,
	// 			  "columnCount": 2
	// 			},
	// 			"tbColor": {
	// 			  "red": 1.0,
	// 			  "green": 0.,
	// 			  "blue": 0.4
	// 			}
	//
	// 		}
	//
	// 	]
	//
}

// func (g GoogleSheets) UpdateGoogleSheetReport(sheetID string, preparedData []PreparedDateToAppendIntoSpreadSheet) {
// 	// var mapSheetsFromDate map[time.Time][]*sheets.Sheet
// 	// var sheetID string = "1Z7KluzcrGpLu_TVyGUBicwt_mpc13FhGlfGluUXh-8"

// // var mapSheet mapGoogleSheetsByMonth
// 	// var pointMapSheet = &mapSheet

// // sheetsService, err := getGoogleSheetSrv()
// 	// if err != nil {
// 	// 	if err != nil
// 	// 		log.Fatalf("Unble to retrieve Sheets client: %v", err)
// 	// 	}
// 	// }

// // resp := sheetsService.Spreadsheets.Get(sheetID)
// 	// var ss *sheets.Spreadsheet
// 	// ss, _ = resp.Do()

// // _ = ss

// //ctx := context.Background()
// 	// How the input data should e interpreted.
// 	valueInputOption := "USER_ENTERED" // TODO: pdate placeholder value.
// 	// The new values to apply to the spreadsheet.
// 	//	rangeData := "sheet1!A1:B3"
// 	//	values := [][]interface{}{{sample_A1", "sample_B1"}, {"sample_A2", "sample_B2"}, {"sample_A3", "sample_A3"}}

// //data := []*sheets.ValueRange{} // TODO: Update placeholder value.

// 	rb := &sheets.BatchUpdateValuesRequest{
// 		ValueInputOption: valueInputOption,
// 		//	Data:             data,

// 	// TODO: Add desired fields of the request body.
// 	}

// // preparedData := Prepare(readerReport)
// 	for _, prep := range preparedData {

// 	rb.Data = append(rb.Data, &sheets.ValueRange{
// 			Range:  prep.Range,
// 			Values: prep.Table,
// 		})
// 	}

// // rangeData = "Лист1!A1:B3"
// 	// rb.Data = append(rb.Data,&sheets.ValueRange{
// 	// 	Range:  rangeData,
// 	// 	Values: values,
// 	// })

// // _, err = sheetsService.Spreadsheets.Values.BatchUpdate(spreadsheetId, rb).Context(ctx).Do()
// 	// if err != nil {
// 	//     log.Fatal(er)
// 	// }
// 	// ft.Println("Done.")

// resp, err := g.SheetsService.Spreadsheets.Values.BatchUpdate(sheetID, rb).Context(g.Ctx).Do()
// 	if err != nil {
// 		log.Fatal(err)
// 	}

// // TODO: Change code below to process the `resp` object:
// 	fmt.Printf("%#v\n", resp)

// // for _, googleSheet := range ss.Sheets {
// 	// 	key, err := getTimeFromTitleSheet(googeSheet.Properties.Title)
// 	// 	if err == nil {
// 	// 		(&mapSheet).Inert(key, googleSheet)
// 	// 		// pointMapSheet.Insert(key, googleSeet)
// 	// 	}
// 	// }

// // // resp_sheets := resp.Get("sheets")
// 	// // fmt.Printf("%#v\n", resp_sheets)
// 	// // if err != nil {
// 	// // 	log.Fatal(err)
// 	// // }

// // return mapSheet
// }

func (g *GoogleSheets) startGoogleSheetSrv() {

	b, err := ioutil.ReadFile(g.clientSecretFile)

	if err != nil {
		log.Fatalf("Unble to read client secret file: %v", err)
	}

	// If modifying these scopes, delete your previously saved token.json.
	// config, err := google.ConfigFromJSON(b, "https://www.googleapis.comauth/spreadsheets.readonly")
	g.Ctx = context.Background()
	// creds, err := google.CredntialsFromJSON(ctx, b, "https://www.googleapis.com/auth/spreadsheets")
	creds, err := google.CredentialsFromJSON(g.Ctx, b, "https://www.googleapis.com/auth/spreadsheets")
	if err != nil {
		log.Fatal(err)
	}

	client := oauth2.NewClient(g.Ctx, creds.TokenSource) //credentials. getClient(config)

	srv, err := sheets.New(client)
	if err != nil {
		log.Fatalf("Unble to retrieve Sheets client: %v", err)
	}

	g.SheetsService = srv
}

func getMonthBefore(t time.Time) time.Time {
	return time.Date(t.Year(), t.Month()-1, 1, 0, 0, 0, 0, time.UTC)
}

func (m mapGoogleSheetsByMonth) getSheetsForDateWithMonthBefore(date time.Time) []*sheets.Sheet {
	sheetsOfDate := m.getSheetsOfDate(date)
	sheetsOfDateBefore := m.getSheetsForDateWithMonthBefore(date)
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

// / func (g *GoogleSheets) getSheetsList(sheetsService *sheets.Service, getTime func(string) (time.Time, error)) mapGoogleSheetsByMonth {
// // 	// True if grid data should be returned.
// // 	// This parameter is ignored if a field ask was set in the request.
// // 	// includeGridData := false // TODO: Update placeholder value.
// // 	// ctx := context.Background()

// / 	// The ranges to retrieve from the spreadsheet.
// // 	// ranges := []string{} // TODO: Update placehoder value.

// / 	//resp, err := sheetsService.Spreadsheets.Get(spreadsheetId).Ranges(ranges...).IncludeGridData(includeGridData).Context(ctx).Do()
// // 	resp := sheetsService.Spreadsheets.Get(g.spreadsheetID)
// // 	var ss *sheets.Spreadsheet
// // 	ss, _ = resp.Do()
// // 	var mapSheet mapGogleSheetsByMonth
// // 	// var pointMapSheet = &mapSheet

// / 	for _, googleSheet := range ss.Sheets {
// // 		key, err := getTime(googleSheet.Properies.Title)
// // 		if err == nil {
// // 			(&mapSheet).Inert(key, googleSheet)
// // 			// pointMapSheet.Insert(key, googleSeet)
// // 		}
// // 	}
// // 	/ // resp_sheets := resp.Get("sheets")
// // 	// // fmt.Printf("%#v\n", resp_sheets)
// // 	// // if err != nil {
// // 	// // 	log.Fatal(err)
// // 	// // }

// / 	return mapSheet
// 	// fmt.Printf("#v\n", len(mapSheet.mapSheets))
// 	// fmt.Printf("%#v\n", mapSheet)

// 	// Prints the names and majors of students in a sample spreadsheet:
// 	// https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUqptlbs74OgvE2upms/edit
// 	// spreadsheetId := "1_EokmBJWtNKYMbQZEeuTrzk7fwLy5z1H-IbzR1v2q14"
// 	// spreadsheetId := "183IDyrxg5PczVLewXMronbduZy50ukDiqNUGgnloqQQ"
// 	// readRange := "МОСГОРТРАНС!A1:E5"
// 	// resp, err := srv.Spreadsheets.Vaues.Get(spreadsheetId, readRange).Do()
// 	// if err != nil {
// 	// 	log.Fatalf("Unble to retrieve data from sheet: %v", err)
// 	// }

// 	// if len(resp.Values) == 0 {
// 	// 	fmt.Println("No data foun.")
// 	// } else {
// 	// 	fmt.Pritln("Name, Major:")
// 	// 	for _, row := range resp.Vaues {
// 	// 		// Print columns A and E, which orrespond to indices 0 and 4.
// 	// 		fmt.Printf("%s\n", row[0])
// 	// 	}
// 	// }

// 	// writeRange := "A1"

//	// var vr sheets.ValueRange

// 	// myval := []interface{}{"One", "Two", "Three"}
// 	// vr.Values = append(vr.Values, myval)

// 	// _, err = srv.Spreadsheets.Values.Update(spreadsheetId, writeRange, &vr).ValueInputOption("RAW").Do()
// 	// if err != nil {
// 	// 	log.Fatalf("Unble to retrieve data from sheet. %v", err)
// 	// }
// }

////Insert Перебираем список считанных из эксель данных и определяем, на каких листах они могут лежать.
// func (t *mapGoogleSheetsByMonth) Insert(month time.Time, gs *sheets.Sheet) {
// 	if t.mapSheets == nil {
// 		t.mapSheets = make(map[time.Time]*sheets.Sheet)
// 	}
// 	googleSheets, ok := t.mapSheets[month]
// 	if ok {
// 		googleheets = append(googleSheets, gs)
// 	} else {
// 		var gooleSheets = make([]*sheets.Sheet, 0, 5)
// 		googleSheets = append(googleSheets, gs)
// 		// fmt.Println("Помещение", month)

// 	}
// 	t.mapSheets[month] = googleSheets
// }

//MakeFilter создаёт именованный фильтр для страницы
func (g *GoogleSheets) MakeFilter(spreadSheetID string, title string, gridRange *sheets.GridRange) {
	filterView := new(sheets.FilterView)

	filterView.Range = gridRange
	req := sheets.Request{
		AddFilterView: &sheets.AddFilterViewRequest{
			//&sheets.SheetProperties{
			Filter: filterView,
			// Title:  tite,
		},
	}

	rb := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{&req},
	}

	resp, err := g.SheetsService.Spreadsheets.BatchUpdate(spreadSheetID, rb).Context(g.Ctx).Do()
	if err != nil {
		log.Fatal(err)
	}
	fmt.Printf("%#v\n", resp)
}

// service = self.service

// # [START sheets_filter_views]
// my_range = {
//     'sheetId': 0,
//     'startRowIndex': 0,
//     'startColumnIndex': 0,
// }
// addFilterViewRequest = {
//     'addFilterView': {
//         'filter': {
//             'title': 'Sample Filter',
//             'range': my_range,
//             'sortSpecs': [{
//                 'dimensionIndex': 3,
//                 'sortOrder': 'DESCENDING'
//             }],
//             'criteria': {
//                 0: {
//                     'hiddenValues': ['Panel']
//                 },
//                 6: {
//                     'condition': {
//                         'type': 'DATE_BEFORE',
//                         'values': {
//                             'userEnteredValue': '4/30/2016'
//                         }
//                     }
//                 }
//             }
//         }
//     }
// }

// body = {'requests': [addFilterViewRequest]}
// addFilterViewResponse = service.spreadsheets() \
//    .batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

// duplicateFilterViewRequest = {
//     'duplicateFilterView': {
//     'filterId':
//         addFilterViewResponse['replies'][0]['addFilterView']['filter']
//             ['filterViewId']
//     }
// }

// body = {'requests': [duplicateFilterViewRequest]}
// duplicateFilterViewResponse = service.spreadsheets() \
//    .batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

// updateFilterViewRequest = {
//     'updateFilterView': {
//         'filter': {
//             'filterViewId': duplicateFilterViewResponse['replies'][0]
//                 ['duplicateFilterView']['filter']['filterViewId'],
//             'title': 'Updated Filter',
//             'criteria': {
//                 0: {},
//                 3: {
//                     'condition': {
//                         'type': 'NUMBER_GREATER',
//                         'values': {
//                             'userEnteredValue': '5'
//                         }
//                     }
//                 }
//             }
//         },
//         'fields': {
//             'paths': ['criteria', 'title']
//         }
//     }
// }

// body = {'requests': [updateFilterViewRequest]}
// updateFilterViewResponse = service.spreadsheets() \
//    .batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
// # [END sheets_filter_views]
