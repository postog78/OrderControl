package excelconnector

import model "OrderControl/app/model"

func WriteDataToGooleSheetFromReaders(listConnector []ConnectorExcelReader) {
	var readerReport []model.TypeReport
	var readerReportFull []model.TypeReport
	for _, reader := range listConnector {
		readerReport = reader.GetData()
		for _, elem := range readerReport {
			readerReportFull = append(readerReportFull, elem)
		}
	}
	model.WriteDataToGoogleSheet(readerReportFull)
}
