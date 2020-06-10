package model

import (
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"path"
	"path/filepath"
	"time"
)

type TypeReport struct {
	NumOrder int
	Weight   int //кг
	Date     time.Time
	Volume   int
}


var spreadsheetId string = "183IDyrxg5PczVLewXMronbduZy50ukDiqNUGgnloqQQ"

func (t TypeReport) String() string {
	return fmt.Sprintf("NumOrder: %v; Weight: %v; Volume: %v Date: %v", t.NumOrder, t.Weight, t.Volume, t.Date)
}

type TypeConnector struct {
	Folder string
	f      func()
}

// Exists reports whether the named file or directory exists.
func FileExists(name string) bool {
	if _, err := os.Stat(name); err != nil {
		if os.IsNotExist(err) {
			return false
		}
	}
	return true
}

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

// rangeDate returns a date range function over start date to end date inclusive.
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
