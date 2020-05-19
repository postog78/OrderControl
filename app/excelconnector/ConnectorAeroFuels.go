package excelconnector

import (
	"errors"
	"fmt"
	"log"
	"time"
	"OrderControl/app/model"

	"github.com/tealeg/xlsx"
)


var excelFileName string = `C:\Users\Dell\Documents\Go\OrderControl\Files\Отчеты базисов об отгрузках\Базис 1\ИНТ_Остатки 2020.xlsx`

func GetDataFromExcelFile() ([]model.TypeReport, error) {
	listReport = make([]model.TypeReport, 0, 20)
	var dateOrder time.Time = time.Date(2020, time.April, 22, 0, 0, 0, 0, time.UTC)
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		fmt.Printf("open failed: %s\n", err)
	}
	numOfSheet := len(xlFile.Sheets)
	switch {
	case numOfSheet == 0:
		return nil, errors.New("This XLSX file contains no sheets.")
		// case sheetIndex >= sheetLen:
		// 	return fmt.Errorf("No sheet %d available, please select a sheet between 0 and %d\n", sheetIndex, sheetLen-1)
	}

	shitName := dateOrder.Format("02.01")
	var currentSheet *xlsx.Sheet
	// fmt.Printf("%v\n", xlFile.Sheet)
	for _, currentSheet = range xlFile.Sheets {
		if shitName == currentSheet.Name {
			fmt.Printf("Sheet Name: %s\n", currentSheet.Name)
			// sheet := sheet
			//currentSheet = currentSheet
			break
		}
	}
	sheetRows := currentSheet.MaxRow
	for i := 5; i < sheetRows; i++ {
		proposalNum, err := currentSheet.Cell(i, 9)
		// proposalNum, err := sheet.Cell(i, 9)
		if err != nil {
			log.Fatal(err)
		}
		proposalDate, err := currentSheet.Cell(i, 10)
		if err != nil {
			log.Fatal(err)
		}
		if proposalNum==nil{
			continue
		}
		var elem TypeReport
		
		elem = TypeReport{}
		append(listReport, elem) 
		fmt.Printf("%v\t%v\n", proposalDate, proposalNum) // Print values in columns B and D
	}

	// for rows.Next() {
	// 	row := rows.Columns()
	// 	fmt.Printf("%s\t%s\n", row[9], row[10]) // Print values in columns B and D
	// }
	// 	for _, row := range sheet.Rows {
	// 		for _, cell := range row.Cells {
	// 			text, _ := cell.String()
	// 			fmt.Printf("%s\n", text)
	// 		}
	// 	}
	// }
	return listReport, nil
}
