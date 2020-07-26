package controller

import (
	"OrderControl/app/excelconnector"
	"OrderControl/app/model"
	"html/template"
	"log"
	"net/http"
	"time"
)

var templates *template.Template

//PageMainView тип, хранящий данные для заглавной страницы
type PageMainView struct {
	Title                       string
	DateBegin, DateEnd, DateNow time.Time
	DayOfWeek                   string
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

	switch now.Weekday() {
	case time.Monday:
		p.DayOfWeek = "Понедельник"
	case time.Tuesday:
		p.DayOfWeek = "Вторник"
	case time.Wednesday:
		p.DayOfWeek = "Среда"
	case time.Thursday:
		p.DayOfWeek = "Четверг"
	case time.Friday:
		p.DayOfWeek = "Пятница"
	case time.Saturday:
		p.DayOfWeek = "Суббота"
	case time.Sunday:
		p.DayOfWeek = "Воскресенье"
	}

	p.DateNow = now
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
	var listConnector []excelconnector.ConnectorExcelReader
	listConnector = append(listConnector, new(excelconnector.AeroFuels))
	listConnector = append(listConnector, new(excelconnector.OkaCenter))
	listConnector = append(listConnector, new(excelconnector.RezervA))
	listConnector = append(listConnector, new(excelconnector.Istra))
	listConnector = append(listConnector, new(excelconnector.IstraOpt))
	listConnector = append(listConnector, new(excelconnector.WhitePillars))
	listConnector = append(listConnector, new(excelconnector.MainSpecialConstruction))

	// var dateBegin time.Time = time.Date(2020, time.April, 20, 0, 0, 0, 0, time.UTC)
	// var dateEnd time.Time = time.Date(2020, time.July, 22, 0, 0, 0, 0, time.UTC)

	for _, reader := range listConnector {
		reader.Init()
		_, err := reader.Read(dateBegin, dateEnd)
		if err != nil {
			log.Fatalf("Не получилось прочитать файл %s. %v", reader.GetName(), err)
		}
	}

	excelconnector.UpdateGoogleSheetCollectedInformation(listConnector)

	// if err := shadow.ShadowControlFromUser(comp, login); err != nil {
	// 	fmt.Fprintln(w, "Ошибка ", err)
	// 	return
	// }

	http.Redirect(w, r, "/result/", http.StatusTemporaryRedirect)
	// fmt.Fprintln(w, "Сгенерированный лист смотрите по адресу (потом здесь будет ссылка, а пока просто надо открыть)")
}
