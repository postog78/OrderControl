package excelconnector

import (
	model "OrderControl/app/model"
	"errors"
	"fmt"
	"log"
	"path"
	"time"

	"github.com/tealeg/xlsx"
)

// AeroFuels - структура, хранящая данные Аэрофьюэлза
type AeroFuels struct {
	Name       string
	PathToDir  string
	ListReport []model.TypeReport
}

// var excelFileName string = `C:\Users\Dell\Documents\Go\OrderControl\Files\Отчеты базисов об отгрузках\Базис 1\ИНТ_Остатки 2020.xlsx`

// Init первоначальная инициализация объекта
func (basis *AeroFuels) Init() {
	basis.Name = "Аэрофьюэлз"
	// basis.PathToDir = "C:\\Users\\User\\go\\src\\OrderControl\\Файлы\\Отгрузки\\Аэрофьюэлз" //path.Join(pathToShipments, "Аэрофьюэл")
	// basis.PathToDir = "C:/Users/User/go/src/OrderControl/Файлы/Отгрузки/Аэрофьюэлз"         //path.Join(pathToShipments, "Аэрофьюэл")
	basis.PathToDir = path.Join(pathToShipments, basis.Name)
}

func (basis *AeroFuels) Read(dateBegin, dateEnd time.Time) ([]model.TypeReport, error) {

	basis.ListReport = nil
	var fullListReport []model.TypeReport = make([]model.TypeReport, 0, 50)

	var dateOrder time.Time //Дата, на которую смотрим загрузку

	for _, excelFileName := range model.GetFiles(basis.PathToDir) {
		xlFile, err := xlsx.OpenFile(excelFileName)
		if err != nil {
			fmt.Printf("open failed: %s\n", err)
			return nil, err
		}

		numOfSheet := len(xlFile.Sheets)
		switch {
		case numOfSheet == 0:
			return nil, errors.New("This XLSX file contains no sheets")
			// case sheetIndex >= sheetLen:
			// 	return fmt.Errorf("No sheet %d available, please select a sheet between 0 and %d\n", sheetIndex, sheetLen-1)
		}

		for _, dateOrder = range model.RangeDate(dateBegin, dateEnd) {
			shitName := dateOrder.Format("02.01")
			var currentSheet *xlsx.Sheet
			// fmt.Printf("%v\n", xlFile.Sheet)
			shitNameWasFound := false
			for _, currentSheet = range xlFile.Sheets {
				if shitName == currentSheet.Name {
					shitNameWasFound = true
					// fmt.Printf("Sheet Name: %s\n", currentSheet.Name)
					// sheet := sheet
					//currentSheet = currentSheet
					break
				}
			}

			if !shitNameWasFound {
				continue
			}

			listReport := make([]model.TypeReport, 0, 20)

			sheetRows := currentSheet.MaxRow
			for i := 5; i < sheetRows; i++ {
				proposalNum, err := currentSheet.Cell(i, 9)
				// proposalNum, err := sheet.Cell(i, 9)
				if err != nil {
					log.Fatal(err)
				}
				if proposalNum == nil || proposalNum.Value == "" {
					continue
				}

				proposalDate, err := currentSheet.Cell(i, 10)
				if err != nil {
					log.Fatal(err)
				}

				proposalWeight, err := currentSheet.Cell(i, 6)
				if err != nil {
					log.Fatal(err)
				}

				proposalVolume, err := currentSheet.Cell(i, 4)
				if err != nil {
					log.Fatal(err)
				}

				numOrder, _ := proposalNum.Int()
				weight, _ := proposalWeight.Int()
				// date, err := strconv.Atoi(proposalNum)
				date, _ := proposalDate.GetTime(proposalDate.Row.Sheet.File.Date1904)
				//date, _ := proposalDate.GetTime(false)
				volume, _ := proposalVolume.Int()

				var elem model.TypeReport

				elem = model.TypeReport{
					NumOrder: numOrder,
					Weight:   weight,
					Date:     date,
					Volume:   volume,
				}
				listReport = append(listReport, elem)
				fmt.Println(basis.GetName(), elem) // Print values in columns B and D
			}
			for _, elem := range listReport {
				fullListReport = append(fullListReport, elem)
			}
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

//GetData возвращает данные ЛистРепорта
func (basis *AeroFuels) GetData() []model.TypeReport {
	return basis.ListReport
}

//GetName возвращает данные имени коннектора
func (basis *AeroFuels) GetName() string {
	return basis.Name
}
