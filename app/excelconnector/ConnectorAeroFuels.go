package excelconnector

import (
	model "OrderControl/app/model"
	"errors"
	"fmt"
	"log"
	"time"

	"github.com/tealeg/xlsx"
)

type AeroFuels struct {
	Name       string
	PathToDir  string
	ListReport []model.TypeReport
}

// var excelFileName string = `C:\Users\Dell\Documents\Go\OrderControl\Files\Отчеты базисов об отгрузках\Базис 1\ИНТ_Остатки 2020.xlsx`

func (basis *AeroFuels) Init() {
	basis.Name = "Аэрофьюэл"
	basis.PathToDir = "C:\\Users\\User\\go\\src\\OrderControl\\Файлы\\Отгрузки\\Аэрофьюэлз" //path.Join(pathToShipments, "Аэрофьюэл")
}

func (basis *AeroFuels) Read(dateBegin, dateEnd time.Time) ([]model.TypeReport, error) {

	basis.ListReport = nil
	var fullListReport []model.TypeReport = make([]model.TypeReport, 0, 50)

	var dateOrder time.Time //Дата, на которую смотрим загрузку

	for _, excelFileName := range model.GetFiles(basis.PathToDir) {
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

		for _, dateOrder = range model.RangeDate(dateBegin, dateEnd) {
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

			listReport := make([]model.TypeReport, 0, 20)

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
				if proposalNum == nil {
					continue
				}
				var elem model.TypeReport

				elem = model.TypeReport{
					NumOrder: 1,
					Weight:   2,
					Date:     time.Now(),
					Volume:   3,
				}
				listReport = append(listReport, elem)
				fmt.Printf("%v\t%v\n", proposalDate, proposalNum) // Print values in columns B and D
			}
			_ = copy(fullListReport, listReport)
		}
	}
	// var dateOrder time.Time =

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
	basis.ListReport = fullListReport
	return fullListReport, nil
}

func (basis *AeroFuels) GetData() []model.TypeReport {
	return basis.ListReport
}

func (basis *AeroFuels) GetName() string {
	return basis.Name
}
