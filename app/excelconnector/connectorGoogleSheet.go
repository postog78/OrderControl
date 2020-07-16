package excelconnector

import (
	model "OrderControl/app/model"
)

//WriteDataToGooleSheetFromReaders записываем данные в гугл файл
// func WriteDataToGooleSheetFromReaders(listConnector []ConnectorExcelReader) {
// 	readerReportFull := getFulllistReaderReport(listConnector)

// 	// var readerReport []model.TypeReport
// 	// var readerReportFull []model.TypeReport
// 	// for _, reader := range listConnector {
// 	// 	readerReport = reader.GetData()
// 	// 	for _, elem := range readerReport {
// 	// 		readerReportFull = append(readerReportFull, elem)
// 	// 	}
// 	// }
// 	model.WriteDataToGoogleSheet(readerReportFull)
// }

//UpdateGoogleSheetCollectedInformation записываем данные в гугл файл
func UpdateGoogleSheetCollectedInformation(listConnector []ConnectorExcelReader) {
	readerReportFull := getFulllistReaderReport(listConnector)

	var ci *model.CollectedInformation
	ci = new(model.CollectedInformation)
	ci.Init(readerReportFull)
	ci.Go()
}

func getFulllistReaderReport(listConnector []ConnectorExcelReader) []model.TypeReport {
	var readerReport []model.TypeReport
	var readerReportFull []model.TypeReport
	for _, reader := range listConnector {
		readerReport = reader.GetData()
		for _, elem := range readerReport {
			readerReportFull = append(readerReportFull, elem)
		}
	}
	return readerReportFull
}
