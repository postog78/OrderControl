package model

import (
	"context"
	"errors"
	"io/ioutil"
	"log"
	"regexp"
	"strconv"
	"time"

	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/sheets/v4"
)

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
	getSheetsList(srv)
	//Перебираем список считанных из эксель данных и определяем, на каких листах они могут лежать.
	//Запоминаем путь в гугл шит для каждой позиции. //Если позиций несколько, то остаётся только первая найденная.
	//Пишем данные
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

func getSheetsList(sheetsService *sheets.Service) {
	// True if grid data should be returned.
	// This parameter is ignored if a field mask was set in the request.
	// includeGridData := false // TODO: Update placeholder value.
	// ctx := context.Background()

	// The ranges to retrieve from the spreadsheet.
	// ranges := []string{} // TODO: Update placeholder value.

	//resp, err := sheetsService.Spreadsheets.Get(spreadsheetId).Ranges(ranges...).IncludeGridData(includeGridData).Context(ctx).Do()
	resp := sheetsService.Spreadsheets.Get(spreadsheetId)
	var ss *sheets.Spreadsheet
	ss, _ = resp.Do()
	map_sheet := make(map[time.Time]*sheets.Sheet)

	for _, google_sheet := range ss.Sheets {
		key, err := getTimeFromTitleSheet(google_sheet.Properties.Title)
		if err == nil {
			map_sheet[key] = google_sheet
		}
	}
	// // resp_sheets := resp.Get("sheets")
	// // fmt.Printf("%#v\n", resp_sheets)
	// // if err != nil {
	// // 	log.Fatal(err)
	// // }

	// fmt.Printf("%#v\n", sheetList)

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

func getTimeFromTitleSheet(title string) (ret_time time.Time, err error) {
	if len(title) == 0 {
		err = errors.New("No title for sheet")
		return
	}
	re := regexp.MustCompile(`(Январь|Февраль|Март|Апрель|Май|Июнь|Июль|Август|Сентябрь|Октябрь|Ноябрь|Декабрь) (\d{4})`)
	map_month := map[string]time.Month{
		"Январь": time.January, "Февраль": time.February, "Март": time.March,
		"Апрель": time.April, "Май": time.May, "Июнь": time.June,
		"Июль": time.July, "Август": time.August, "Сентябрь": time.September,
		"Октябрь": time.October, "Ноябрь": time.November, "Декабрь": time.December,
	}
	//На всякий случай берём на три года назад и только с целью разработки
	var matches [][]string = re.FindAllStringSubmatch(title, -1)
	if !(len(matches) == 1 && len(matches[0]) == 3) {
		err = errors.New("This is not a template")
		return
	}
	year, _ := strconv.Atoi(matches[0][2])
	ret_time = time.Date(year, map_month[matches[0][1]], 0, 0, 0, 0, 0, time.UTC)
	return
	// matches[1]
	// Продажа Февраль 2020 (РЕГИОНЫ)
	// ВЛПДС ПРОДАЖА (Январь 2020)
	// Загрузка_марийка
	// Продажа Декабрь 2019 (РЕГИОНЫ)

}
