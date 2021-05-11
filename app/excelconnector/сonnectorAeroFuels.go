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

// AeroFuels - структура, хранящая данные Аэрофьюэлза
type AeroFuels struct {
	Name             string
	PathToDir        string
	ListReport       []model.TypeReport
	colTypeOfProduct int
	colDriver        int
	rowInfo          int
}

func (basis *AeroFuels) ItIsMyFormat(excelFileName string) (bool, error) {

	strCaption := `Реестр выданных ТТН ООО "АЭРОФЬЮЭЛЗ Нижний Новгород"`
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		fmt.Printf("open failed: %s\n", err)
		return false, err
	}

	numOfSheet := len(xlFile.Sheets)
	switch {
	case numOfSheet == 0:
		return false, errors.New("This XLSX file contains no sheets")
		// case sheetIndex >= sheetLen:
		// 	return fmt.Errorf("No sheet %d available, please select a sheet between 0 and %d\n", sheetIndex, sheetLen-1)
	}

	myFormat := false

	// listDatesForLoad := list.New()
	// for _, dateOrder = range model.RangeDate(dateBegin, dateEnd) {
	// 	shitName := dateOrder.Format("02.01")
	// 	var currentSheet *xlsx.Sheet
	// 	// fmt.Printf("%v\n", xlFile.Sheet)
	// 	shitNameWasFound := false

	for _, currentSheet := range xlFile.Sheets {

		if !myFormat {
			valueInfo, err := currentSheet.Cell(basis.rowInfo, 0)
			if err != nil {
				log.Fatal(err)
			}
			tmpVal, _ := currentSheet.Cell(2, 0)
			log.Println(tmpVal.String())

			tmpVal, _ = currentSheet.Cell(2, 1)
			log.Println(tmpVal.String())

			tmpVal, _ = currentSheet.Cell(3, 1)
			log.Println(tmpVal.String())

			if strings.Index(strings.ToUpper(valueInfo.String()), strings.ToUpper(strCaption)) > -1 {
				myFormat = true
			}

		}
		// listDatesForLoad.PushFront(dateSheet)
	}

	return myFormat, nil
}

// var excelFileName string = `C:\Users\Dell\Documents\Go\OrderControl\Files\Отчеты базисов об отгрузках\Базис 1\ИНТ_Остатки 2020.xlsx`

// Init первоначальная инициализация объекта
func (basis *AeroFuels) Init() {
	basis.Name = "Аэрофьюэлз"
	// basis.PathToDir = "C:\\Users\\User\\go\\src\\OrderControl\\Файлы\\Отгрузки\\Аэрофьюэлз" //path.Join(pathToShipments, "Аэрофьюэл")
	// basis.PathToDir = "C:/Users/User/go/src/OrderControl/Файлы/Отгрузки/Аэрофьюэлз"         //path.Join(pathToShipments, "Аэрофьюэл")
	basis.PathToDir = path.Join(pathToShipments, basis.Name)
	basis.colTypeOfProduct = 3
	basis.colDriver = 11
	basis.rowInfo = 2
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

				proposalComment, err := currentSheet.Cell(i, 11)
				if err != nil {
					log.Fatal(err)
				}

				proposalTypeOfProduct, err := currentSheet.Cell(i, basis.colTypeOfProduct)
				if err != nil {
					log.Fatal(err)
				}

				// proposalDate, err := currentSheet.Cell(i, 10)
				// if err != nil {
				// 	log.Fatal(err)
				// }

				proposalWeight, err := currentSheet.Cell(i, 6)
				if err != nil {
					log.Fatal(err)
				}

				proposalVolume, err := currentSheet.Cell(i, 4)
				if err != nil {
					log.Fatal(err)
				}

				var driver string

				if basis.colDriver >= 0 {
					proposalDriver, err := currentSheet.Cell(i, basis.colDriver)
					if err != nil {
						log.Fatal(err)
					}
					driver = proposalDriver.String()
				}

				numOrder, _ := proposalNum.Int()
				weight, _ := proposalWeight.Int()
				// date, _ := proposalDate.GetTime(proposalDate.Row.Sheet.File.Date1904)
				date := dateOrder
				volume, _ := proposalVolume.Int()
				typeOfProduct := proposalTypeOfProduct.String()

				var elem model.TypeReport

				elem = model.TypeReport{
					NumOrder:      numOrder,
					Weight:        weight,
					Date:          date,
					Volume:        volume,
					Driver:        driver,
					BasisName:     basis.GetName(),
					SheetName:     currentSheet.Name,
					Row:           i + 1,
					Comment:       proposalComment.String(),
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

//GetData возвращает данные ЛистРепорта
func (basis *AeroFuels) GetData() []model.TypeReport {
	return basis.ListReport
}

//GetName возвращает данные имени коннектора
func (basis *AeroFuels) GetName() string {
	return basis.Name
}
