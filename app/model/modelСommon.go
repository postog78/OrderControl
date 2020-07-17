package model

import (
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"path"
	"path/filepath"
	"time"

	"google.golang.org/api/sheets/v4"
)

//TypeReport - структура, содержащая основные данные по ответу нефтебазы по отгрузке
type TypeReport struct {
	NumOrder  int
	Weight    int //кг
	Date      time.Time
	Volume    int
	Comment   string
	BasisName string
	SheetName string
	Row       int
	FileName string
}

type dataToSearchInGoogleSheets struct {
	rep    TypeReport
	sheets []*sheets.Sheet
}

var spreadsheetID string = "183IDyrxg5PczVLewXMronbduZy50ukDiqNUGgnloqQQ"

func (t TypeReport) String() string {
	return fmt.Sprintf("Basis: %s; Sheet: %s; Row: %d, NumOrder: %v; Weight: %v; Volume: %v Date: %v, Comment: %s",
		t.BasisName,
		t.SheetName,
		t.Row,
		t.NumOrder,
		t.Weight,
		t.Volume,
		t.Date,
		t.Comment,
	)
}

// //TypeConnector структура, содержащая путь к папке и функцию, которую будут
// type TypeConnector struct {
// 	Folder string
// 	f      func()
// }

//FileExists reports whether the named file or directory exists.
func FileExists(name string) bool {
	if _, err := os.Stat(name); err != nil {
		if os.IsNotExist(err) {
			return false
		}
	}
	return true
}

//GetFiles получает все документы XLSX из папки, переданной как параметр
func GetFiles(nameDir string) (fileNameList []string) {
	if !FileExists(nameDir) {
		return
	}

	files, err := ioutil.ReadDir(nameDir)
	if err != nil {
		log.Fatal(err)
	}

	for _, file := range files {
		if !file.IsDir() && filepath.Ext(file.Name()) == ".xlsx" {
			fileNameList = append(fileNameList, path.Join(nameDir, file.Name()))
		}
		// fmt.Println(file.Name(), file.IsDir())
	}
	return
}

// RangeDate returns a date range function over start date to end date inclusive.
// After the end of the range, the range function returns a zero date,
// date.IsZero() is true.
func RangeDate(start, end time.Time) (list []time.Time) {
	// y, m, d := start.Date()
	// start = time.Date(y, m, d, 0, 0, 0, 0, time.UTC)
	// y, m, d = end.Date()
	// end = time.Date(y, m, d, 0, 0, 0, 0, time.UTC)

	// return func() time.Time {
	//     if start.After(end) {
	//         return time.Time{}
	//     }
	//     date := start
	//     start = start.AddDate(0, 0, 1)
	//     return date
	// }
	for d := start; d.After(end) == false; d = d.AddDate(0, 0, 1) {
		list = append(list, d)
		// fmt.Println(d.Format("2006-01-02"))
	}
	return
}

// type GoogleSheets struct {
// 	Sheets []*sheets.Sheet
// }


type columnNum int



type sheetSetting struct {
	NumOrder columnNum //column
	Weight   columnNum //column
	Date     columnNum //column
}
