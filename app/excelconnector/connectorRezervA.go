package excelconnector

import (
	model "OrderControl/app/model"
	"errors"
	"fmt"
	"log"
	"path"
	"strings"
	"time"

	"github.com/tealeg/xlsx"
)

//RezervA коннектор Резерв А
type RezervA struct {
	Name                               string
	PathToDir                          string
	ListReport                         []model.TypeReport
	colComment, colVolume, colWeight   int
	colDate                            int //Особенная колонка, одно значение на лист
	rowDate                            int //Особенная строка, одно значение на лист
	colSummary                         int //Колонка, где в первой колонке стоит "итого"
	colTypeOfProduct, rowTypeOfProduct int
}

// var excelFileName string = `C:\Users\Dell\Documents\Go\OrderControl\Files\Отчеты базисов об отгрузках\Базис 1\ИНТ_Остатки 2020.xlsx`

//Init инициализация коннектора к экселю
func (basis *RezervA) Init() {
	basis.Name = "Резерв А"
	basis.PathToDir = path.Join(pathToShipments, "РезервА")
	basis.colComment = 5
	basis.colVolume = 6
	basis.colWeight = 9
	basis.colDate = 0
	basis.rowDate = 1
	basis.colTypeOfProduct = 0
	basis.rowTypeOfProduct = 0
}

func (basis *RezervA) Read(dateBegin, dateEnd time.Time) ([]model.TypeReport, error) {

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

			listReport := make([]model.TypeReport, 0, 20)

			//Общая дата на весь лист
			proposalDate, err := currentSheet.Cell(basis.rowDate, basis.colDate)
			if err != nil {
				log.Fatal(err)
			}

			//Общее наименование продукта
			proposalTypeOfProduct, err := currentSheet.Cell(basis.rowTypeOfProduct, basis.colTypeOfProduct)
			if err != nil {
				log.Fatal(err)
			}

			sheetRows := currentSheet.MaxRow
			for i := 6; i < sheetRows; i++ {

				//Проверяем, что ещё не добрались до строки итого
				proposalSummary, _ := currentSheet.Cell(i, basis.colSummary)
				if strings.TrimSpace(strings.ToLower(proposalSummary.String())) == "итого" {
					break
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

				weight, _ := proposalWeight.Int()
				if weight <= 0 {
					continue
				}
				date, _ := basis.getAbnormalDateFromString(proposalDate.String())
				dateInPeriod := date.Equal(dateBegin) || date.Equal(dateEnd) || (date.After(dateBegin) && date.Before(dateEnd))
				if !dateInPeriod {
					continue
				}
				volume, _ := proposalVolume.Int()
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
				//fmt.Println(elem) // Print values in columns B and D
			}
			for _, elem := range listReport {
				fullListReport = append(fullListReport, elem)
			}
		}
	}

	basis.ListReport = fullListReport
	return fullListReport, nil
}

//GetData выдача списка прочитанных данных из эксель
func (basis *RezervA) getAbnormalDateFromString(str string) (date time.Time, err error) {
	words := strings.Split(str, " ")
	if len(words) == 0 {
		return date, nil
	}
	return time.Parse(`02,01,2006`, words[0])
}

//GetData выдача списка прочитанных данных из эксель
func (basis *RezervA) GetData() []model.TypeReport {
	return basis.ListReport
}

//GetName выдача данных по имени коннектора
func (basis *RezervA) GetName() string {
	return basis.Name
}
