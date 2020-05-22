package excelconnector

import model "OrderControl/app/model"

func WriteDataToGooleSheetFromReaders(listConnector []ConnectorExcelReader) {
	var readerReport []model.TypeReport
	var readerReportFull []model.TypeReport
	for _, reader := range listConnector {
		readerReport = reader.GetData()
		_ = copy(readerReportFull, readerReport)
	}
	model.WriteDataToGooleSheet(readerReportFull)
}
