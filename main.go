package main

import (
	"OrderControl/app/excelconnector"
	"log"
	"time"
)

func main() {

	//Нам известно о следующих форматах файлов
	//Аэрофьюэлз
	// var listConnector []model.TypeConnector = [
	// 		{"Аэрофьюэлз", excelconnector.GetDataFromExcelFileAeroFuels()}
	// 	]
	var listConnector []excelconnector.ConnectorExcelReader
	//var aeroFuels excelconnector.AeroFuels
	//var reader excelconnector.ConnectorExcelReader = &aeroFuels
	listConnector = append(listConnector, new(excelconnector.AeroFuels))
	listConnector = append(listConnector, new(excelconnector.OkaCenter))
	listConnector = append(listConnector, new(excelconnector.RezervA))
	listConnector = append(listConnector, new(excelconnector.Istra))
	listConnector = append(listConnector, new(excelconnector.IstraOpt))

	var dateBegin time.Time = time.Date(2020, time.April, 20, 0, 0, 0, 0, time.UTC)
	var dateEnd time.Time = time.Date(2020, time.July, 22, 0, 0, 0, 0, time.UTC)

	for _, reader := range listConnector {
		reader.Init()
		_, err := reader.Read(dateBegin, dateEnd)
		if err != nil {
			log.Fatalf("Не получилось прочитать файл %s. %v", reader.GetName(), err)
		}
	}

	excelconnector.UpdateGoogleSheetCollectedInformation(listConnector)
	//excelconnector.WriteDataToGooleSheetFromReaders(listConnector)
	// li{excelconnector.GetDataFromExcelFileAeroFuels()}

	// var listReport []model.TypeReport
	// for connector
	// listReport, _ = excelconnector.GetDataFromExcelFile()
	// if len(listReport) > 0 {
	// 	return
	// }
}
