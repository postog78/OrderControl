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

//FlatTable коннектор Истры
type FlatTable struct {
	Name                             string
	PathToDir                        string
	ListReport                       []model.TypeReport
	colComment, colVolume, colWeight int
	colDate                          int //Особенная колонка, одно значение на лист
	colTypeOfProduct                 int
	rowBegin                         int
	weightInTone                     bool
}

// var excelFileName string = `C:\Users\Dell\Documents\Go\OrderControl\Files\Отчеты базисов об отгрузках\Базис 1\ИНТ_Остатки 2020.xlsx`

//Init инициализация коннектора к экселю
func (basis *FlatTable) Init() {
	basis.Name = ""
	basis.PathToDir = path.Join(pathToShipments, "")
	basis.colComment = 8
	basis.colVolume = 13
	basis.colWeight = 12
	basis.colDate = 1
	basis.colTypeOfProduct = 10
	basis.rowBegin = 6
	basis.weightInTone = false
}

func (basis *FlatTable) Read(dateBegin, dateEnd time.Time) ([]model.TypeReport, error) {

	// fmt.Println(dateBegin)
	// fmt.Println(dateEnd)

	basis.ListReport = nil
	var fullListReport []model.TypeReport = make([]model.TypeReport, 0, 50)

	//var dateOrder time.Time //Дата, на которую смотрим загрузку

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

		//for _, dateOrder = range model.RangeDate(dateBegin, dateEnd) {
		// shitName := dateOrder.Format("02.01")
		var currentSheet *xlsx.Sheet
		// fmt.Printf("%v\n", xlFile.Sheet)

		for _, currentSheet = range xlFile.Sheets {

			if currentSheet.Hidden {
				continue
			}

			listReport := make([]model.TypeReport, 0, 20)

			sheetRows := currentSheet.MaxRow
			for i := basis.rowBegin; i < sheetRows; i++ {

				proposalDate, err := currentSheet.Cell(i, basis.colDate)
				if err != nil {
					log.Fatal(err)
				}
				//Если это не дата, то значит это или пропущенная строка или это заголовок
				date, err := proposalDate.GetTime(xlFile.Date1904)
				if err != nil {
					date, err = basis.getDateFromString(proposalDate.String())
					if err != nil {
						continue
					}
				}

				//Проверяем, что ещё не добрались до строки итого
				proposalTypeOfProduct, err := currentSheet.Cell(i, basis.colTypeOfProduct)
				if err != nil {
					log.Fatal(err)
				}

				proposalComment, err := currentSheet.Cell(i, basis.colComment)
				// proposalNum, err := sheet.Cell(i, 9)
				if err != nil {
					log.Fatal(err)
				}
				// if proposalComment == nil || proposalComment.Value == "" {
				// 	continue
				// }

				proposalWeight, err := currentSheet.Cell(i, basis.colWeight)
				if err != nil {
					log.Fatal(err)
				}

				proposalVolume, err := currentSheet.Cell(i, basis.colVolume)
				if err != nil {
					log.Fatal(err)
				}

				comment := proposalComment.String()

				var weight int
				if basis.weightInTone {
					weightTone, err := proposalWeight.Float()
					if err != nil {
						continue
					}
					weight = int(weightTone * 1000)
				} else {
					weight, err = proposalWeight.Int()
					if err != nil {
						continue
					}
				}

				dateInPeriod := date.Equal(dateBegin) || date.Equal(dateEnd) || (date.After(dateBegin) && date.Before(dateEnd))
				if !dateInPeriod {
					continue
				}
				volume, err := proposalVolume.Int()
				if err != nil {
					continue
				}
				typeOfProduct := proposalTypeOfProduct.String()

				var elem model.TypeReport

				elem = model.TypeReport{
					NumOrder:      -1,
					Weight:        weight,
					Date:          date,
					Volume:        volume,
					Comment:       comment,
					BasisName:     basis.GetName(),
					SheetName:     currentSheet.Name,
					Row:           i + 1,
					FileName:      excelFileName,
					TypeOfProduct: typeOfProduct,
				}
				listReport = append(listReport, elem)
				// fmt.Println(elem) // Print values in columns B and D
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

//getDateFromString Получение даты из строки
func (basis *FlatTable) getDateFromString(str string) (date time.Time, err error) {
	return time.Parse(`02.01.2006`, str)
}

//GetData выдача списка прочитанных данных из эксель
func (basis *FlatTable) GetData() []model.TypeReport {
	return basis.ListReport
}

//GetName выдача данных по имени коннектора
func (basis *FlatTable) GetName() string {
	return basis.Name
}
