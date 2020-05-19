package main

import (
	"./app/excelConnector"
	"./app/model"
)

func main() {
	// b, err := ioutil.ReadFile("diversityshipments-31d3872ee94b.json")

	// if err != nil {
	// 	log.Fatalf("Unable to read client secret file: %v", err)
	// }

	// // If modifying these scopes, delete your previously saved token.json.
	// // config, err := google.ConfigFromJSON(b, "https://www.googleapis.com/auth/spreadsheets.readonly")
	// ctx := context.Background()
	// // creds, err := google.CredentialsFromJSON(ctx, b, "https://www.googleapis.com/auth/spreadsheets")
	// creds, err := google.CredentialsFromJSON(ctx, b, "https://www.googleapis.com/auth/spreadsheets")
	// if err != nil {
	// 	log.Fatal(err)
	// }

	// client := oauth2.NewClient(ctx, creds.TokenSource) //credentials. getClient(config)

	// srv, err := sheets.New(client)
	// if err != nil {
	// 	log.Fatalf("Unable to retrieve Sheets client: %v", err)
	// }

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
	var listReport []model.TypeReport
	listReport, _ = excelConnector.GetDataFromExcelFile()
}

// Просто ключ AIzaSyBg2jgYgXgjlhqTJNK3iN3KVeNE9acf0vU
// Идентификатор клиента 510713154079-5oan62fqccksj647mt12v025rdo3jle1.apps.googleusercontent.com
// Секретный код клиента 7O2XOJbj8GGR3oXtVnlU5pH5
