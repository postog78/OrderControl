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

//OkaCenter коннектор Ока Центр
type OkaCenter struct {
	Name       string
	PathToDir  string
	ListReport []model.TypeReport
}

// var excelFileName string = `C:\Users\Dell\Documents\Go\OrderControl\Files\Отчеты базисов об отгрузках\Базис 1\ИНТ_Остатки 2020.xlsx`

//Init инициализация коннектора к экселю
func (basis *OkaCenter) Init() {
	basis.Name = "Ока центр"
	// basis.PathToDir = "C:\\Users\\User\\go\\src\\OrderControl\\Файлы\\Отгрузки\\Аэрофьюэлз" //path.Join(pathToShipments, "Аэрофьюэл")
	// basis.PathToDir = "C:/Users/User/go/src/OrderControl/Файлы/Отгрузки/Аэрофьюэлз"         //path.Join(pathToShipments, "Аэрофьюэл")
	basis.PathToDir = path.Join(pathToShipments, "ОкаЦентр")
}

func (basis *OkaCenter) Read(dateBegin, dateEnd time.Time) ([]model.TypeReport, error) {

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
		currentSheet = xlFile.Sheets[0]

		listReport := make([]model.TypeReport, 0, 20)

		sheetRows := currentSheet.MaxRow
		for i := 5; i < sheetRows; i++ {
			proposalComment, err := currentSheet.Cell(i, 12)
			// proposalNum, err := sheet.Cell(i, 9)
			if err != nil {
				log.Fatal(err)
			}
			if proposalComment == nil || proposalComment.Value == "" {
				continue
			}

			proposalDate, err := currentSheet.Cell(i, 0)
			if err != nil {
				log.Fatal(err)
			}

			proposalWeight, err := currentSheet.Cell(i, 6)
			if err != nil {
				log.Fatal(err)
			}

			proposalVolume, err := currentSheet.Cell(i, 8)
			if err != nil {
				log.Fatal(err)
			}

			comment := proposalComment.String()
			numOrder, err := getNumFromComment(comment)
			if err != nil {
				continue
			}
			weightTone, _ := proposalWeight.Float()
			weight := int(weightTone * 1000)
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

//GetData выдача списка прочитанных данных из эксель
func (basis *OkaCenter) GetData() []model.TypeReport {
	return basis.ListReport
}

//GetName выдача данных по имени коннектора
func (basis *OkaCenter) GetName() string {
	return basis.Name
}

func getNumFromComment(comment string) (int, error) {
	return 1, nil
}
