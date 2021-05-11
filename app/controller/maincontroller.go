package controller

import (
	"OrderControl/app/excelconnector"
	"OrderControl/app/model"
	"fmt"
	"html/template"
	"io"
	"io/ioutil"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"strings"
	"time"

	// xls "github.com/extrame/ole2"
	xls "github.com/extrame/xls"
)

var templates *template.Template
var listConnector []excelconnector.ConnectorExcelReader

//PageMainView тип, хранящий данные для заглавной страницы
type PageMainView struct {
	Title                       string
	DateBegin, DateEnd, DateNow time.Time
	DayOfWeek                   string
	EraseAnotherLists           bool
}

//PageXLSXFileInfoMainView тип, хранящий данные для заглавной страницы
type PageXLSXFileInfoMainView struct {
	Title                       string
	DateBegin, DateEnd, DateNow time.Time
	DayOfWeek                   string
	ConnectorName               string
}

//CreateTemplates Функция, которая компилирует шаблоны приложения в переменную templates
func CreateTemplates() {
	var err error
	if templates, err = template.New("form_view.html").ParseFiles(`app\controller\form_view.html`); err != nil {
		panic(err)
	}
	if templates, err = templates.New("form_result.html").ParseFiles(`app\controller\form_result.html`); err != nil {
		panic(err)
	}
	if templates, err = templates.New("form_getFileFormat.html").ParseFiles(`app\controller\form_getFileFormat.html`); err != nil {
		panic(err)
	}
	if templates, err = templates.New("form_notNeededFile.html").ParseFiles(`app\controller\form_notNeededFile.html`); err != nil {
		panic(err)
	}

}

//InitListConnector создаёт список коннекторов
func InitListConnector() {
	listConnector = append(listConnector, new(excelconnector.AeroFuels))
	listConnector = append(listConnector, new(excelconnector.OkaCenter))
	listConnector = append(listConnector, new(excelconnector.RezervA))
	listConnector = append(listConnector, new(excelconnector.Istra))
	listConnector = append(listConnector, new(excelconnector.IstraOpt))
	listConnector = append(listConnector, new(excelconnector.WhitePillars))
	listConnector = append(listConnector, new(excelconnector.MainSpecialConstruction))
}

//LoadViewPage возвращает страницу со списком пользователей, где есть кнопки, чтобы подключиться к
//пользователям
func LoadViewPage() (p PageMainView, err error) {
	p.Title = "Запуск считывания отгрузок из файлов базисов в формате *.xlsx и вывод результата в Google таблицу"

	now := time.Now()
	var dateBegin time.Time
	dateEnd := now

	if now.Weekday() == time.Monday {
		dateBegin = time.Date(now.Year(), now.Month(), now.Day()-3, 0, 0, 0, 0, time.UTC)
	} else {
		dateBegin = time.Date(now.Year(), now.Month(), now.Day()-1, 0, 0, 0, 0, time.UTC)
	}
	p.DateBegin = model.StartToTheDay(dateBegin)

	p.DateEnd = dateEnd
	p.DayOfWeek = getRusDayOfWeek(now)
	p.DateNow = now
	p.EraseAnotherLists = true

	return
}

func getRusDayOfWeek(now time.Time) (dayOfWeek string) {
	switch now.Weekday() {
	case time.Monday:
		dayOfWeek = "Понедельник"
	case time.Tuesday:
		dayOfWeek = "Вторник"
	case time.Wednesday:
		dayOfWeek = "Среда"
	case time.Thursday:
		dayOfWeek = "Четверг"
	case time.Friday:
		dayOfWeek = "Пятница"
	case time.Saturday:
		dayOfWeek = "Суббота"
	case time.Sunday:
		dayOfWeek = "Воскресенье"
	}
	return
}

//RenderTemplateView получает из мэйна запрос и возвращает страницу
func RenderTemplateView(w http.ResponseWriter, data PageMainView) {
	err := templates.ExecuteTemplate(w, "form_view.html", data)

	if err != nil {
		http.Error(w, err.Error(),
			http.StatusInternalServerError)
	}
}

//RenderTemplateResult получает из мэйна запрос и возвращает страницу
func RenderTemplateResult(w http.ResponseWriter, a string) {
	err := templates.ExecuteTemplate(w, "form_result.html", a)

	if err != nil {
		http.Error(w, err.Error(),
			http.StatusInternalServerError)
	}
}

//RenderTemplateGetFileFormat  получает запрос и возвращает страницу
func RenderTemplateGetFileFormat(w http.ResponseWriter, data PageXLSXFileInfoMainView) {

	now := time.Now()
	var dateBegin time.Time
	dateEnd := now

	if now.Weekday() == time.Monday {
		dateBegin = time.Date(now.Year(), now.Month(), now.Day()-3, 0, 0, 0, 0, time.UTC)
	} else {
		dateBegin = time.Date(now.Year(), now.Month(), now.Day()-1, 0, 0, 0, 0, time.UTC)
	}
	data.DateBegin = model.StartToTheDay(dateBegin)

	data.DateEnd = dateEnd
	data.DayOfWeek = getRusDayOfWeek(now)
	data.DateNow = now

	err := templates.ExecuteTemplate(w, "form_getFileFormat.html", data)

	if err != nil {
		http.Error(w, err.Error(),
			http.StatusInternalServerError)
	}
}

//RenderNotNeededFile  возвращает страницу с описанием, что формат такого файла не поддерживается
func RenderNotNeededFile(w http.ResponseWriter) {
	err := templates.ExecuteTemplate(w, "form_notNeededFile.html", nil)

	if err != nil {
		http.Error(w, err.Error(),
			http.StatusInternalServerError)
	}
}

//UploadFiles сохраняет файлы, которые передал пользователь
func UploadFiles(w http.ResponseWriter, r *http.Request) {
	m, err := r.MultipartReader()
	if err != nil {
		fmt.Println(err)
	}

	var tempFile *os.File

	for {
		part, err := m.NextPart()
		if err == io.EOF {
			break
		}

		if part.FileName() == "" {
			continue
		}

		fileName := part.FileName()
		ext := filepath.Ext(fileName)
		if !strings.EqualFold(strings.ToUpper(ext), ".XLSX") {
			RenderNotNeededFile(w)
			break
		}

		tempFile, err = ioutil.TempFile(os.TempDir(), "spa*"+ext)
		if err != nil {
			fmt.Println(err)
		}
		defer tempFile.Close()

		_, err = io.Copy(tempFile, part)
		if err != nil {
			http.Error(w, err.Error(), http.StatusInternalServerError)
			return
		}

		p, _ := GetConnectorTypeXLSX(tempFile)
		RenderTemplateGetFileFormat(w, p)

	}

	// r.ParseMultipartForm(32 << 20)
	//    file, handler, err := r.FormFile("files")
	//    if err != nil {
	//        fmt.Println(err)
	//        return
	//    }
	//    defer file.Close()
	//    fmt.Fprintf(w, "%v", handler.Header)
	//    dst, err := os.Create("./test/" + part.FileName())
	//    if err != nil {
	// 	   fmt.Println(err)
	//    }
	//    io.Copy(dst, part)
	//    defer f.Close()
	//    io.Copy(f, file)
}

//GetConnectorTypeXLSX - выдаёт коннектор, который подходит под файл
func GetConnectorTypeXLSX(file *os.File) (p PageXLSXFileInfoMainView, err error) {
	excelFileName := file.Name()
	// p.ConnectorName = excelFileName

	if xlFile, err := xls.Open(excelFileName, "utf-8"); err == nil {
		fmt.Println(xlFile.Author)
	}

	for _, reader := range listConnector {
		reader.Init()
		if isFormat, err := reader.ItIsMyFormat(excelFileName); err != nil && isFormat {
			p.ConnectorName = reader.GetName()
			break
		}
	}

	return p, nil
	// xlFile, err := xlsx.OpenFile(excelFileName)
	// if err != nil {
	// 	fmt.Printf("open failed: %s\n", err)
	// 	return p, err
	// }
	// _= xlFile

	// numOfSheet := len(xlFile.Sheets)
	// switch {
	// case numOfSheet == 0:
	// 	return nil, errors.New("This XLSX file contains no sheets")
	// 	// case sheetIndex >= sheetLen:
	// 	// 	return fmt.Errorf("No sheet %d available, please select a sheet between 0 and %d\n", sheetIndex, sheetLen-1)
	// }

	// for _, dateOrder = range model.RangeDate(dateBegin, dateEnd) {
	// 	shitName := dateOrder.Format("02.01")
	// 	var currentSheet *xlsx.Sheet
	// 	// fmt.Printf("%v\n", xlFile.Sheet)
	// 	shitNameWasFound := false
	// 	for _, currentSheet = range xlFile.Sheets {
	// 		if shitName == currentSheet.Name {
	// 			shitNameWasFound = true
	// 			// fmt.Printf("Sheet Name: %s\n", currentSheet.Name)
	// 			// sheet := sheet
	// 			//currentSheet = currentSheet
	// 			break
	// 		}
	// 	}

	// 	if !shitNameWasFound {
	// 		continue
	// 	}
	// }
	return p, err
}

//RunProcess запускает чтение из файлов и запись в гугл таблицу
func RunProcess(w http.ResponseWriter, r *http.Request) {

	if err := r.ParseForm(); err != nil {
		http.Error(w, err.Error(),
			http.StatusInternalServerError)
	}
	dateBeginString := r.FormValue("DateBegin")
	dateEndString := r.FormValue("DateEnd")

	now := time.Now()
	// fmt.Println(now)

	dateBegin, err := time.Parse("2006-01-02", dateBeginString)
	if err != nil {
		dateBegin = model.StartToTheDay(now)
	}
	dateEnd, err := time.Parse("2006-01-02", dateEndString)
	if err != nil {
		dateEnd = now
	}

	// fmt.Println(dateBegin)
	// fmt.Println(dateEnd)

	//Нам известно о следующих форматах файлов

	// var dateBegin time.Time = time.Date(2020, time.April, 20, 0, 0, 0, 0, time.UTC)
	// var dateEnd time.Time = time.Date(2020, time.July, 22, 0, 0, 0, 0, time.UTC)

	for _, reader := range listConnector {
		reader.Init()
		_, err := reader.Read(dateBegin, dateEnd)
		if err != nil {
			log.Fatalf("Не получилось прочитать файл %s. %v", reader.GetName(), err)
		}
	}

	sheetID := excelconnector.UpdateGoogleSheetCollectedInformation(listConnector)

	deleteAllSheet := r.FormValue("EraseAnotherLists")
	if deleteAllSheet == "true" {
		excelconnector.DeleteAllSheetsCollectedInformation(sheetID)
	}

	// if err := shadow.ShadowControlFromUser(comp, login); err != nil {
	// 	fmt.Fprintln(w, "Ошибка ", err)
	// 	return
	// }

	http.Redirect(w, r, "/result/", http.StatusTemporaryRedirect)
	// fmt.Fprintln(w, "Сгенерированный лист смотрите по адресу (потом здесь будет ссылка, а пока просто надо открыть)")
}
