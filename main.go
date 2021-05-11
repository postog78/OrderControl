package main

import (
	// "OrderControl/app/excelconnector"
	controller "OrderControl/app/controller"
	"log"
	"net/http"
	"os/exec"
	"regexp"
	"runtime"
	// "time"
)

func main() {

	controller.CreateTemplates()
	controller.InitListConnector()

	http.HandleFunc("/view/", makeHandler(viewHandler))
	http.HandleFunc("/run/", makeHandler(runHandler))
	http.HandleFunc("/result/", makeHandler(resultHandler))
	http.HandleFunc("/getFileFormat/", makeHandler(getFileFormatHandler))
	// http.HandleFunc("/edit/", makeHandler(editHandler))
	// http.HandleFunc("/save/", makeHandler(saveHandler))
	// http.HandleFunc("/add/", makeHandler(addHandler))
	// http.HandleFunc("/delete/", makeHandler(deleteHandler))

	go open("http://localhost:8081/view/")
	if err := http.ListenAndServe(":8081", nil); err != nil {
		log.Fatal(err)
	}

	//excelconnector.WriteDataToGooleSheetFromReaders(listConnector)
	// li{excelconnector.GetDataFromExcelFileAeroFuels()}

	// var listReport []model.TypeReport
	// for connector
	// listReport, _ = excelconnector.GetDataFromExcelFile()
	// if len(listReport) > 0 {
	// 	return
	// }
}

// open opens the specified URL in the default browser of the user.
func open(url string) error {
	var cmd string
	var args []string

	switch runtime.GOOS {
	case "windows":
		cmd = "cmd"
		args = []string{"/c", "start"}
	case "darwin":
		cmd = "open"
	default: // "linux", "freebsd", "openbsd", "netbsd"
		cmd = "xdg-open"
	}
	args = append(args, url)
	return exec.Command(cmd, args...).Start()
}

var validPath = regexp.MustCompile(
	"^/(view)/$|^/(run)/$|^/(result)/$|^/(getFileFormat)/$")

// "^/(view)/$|^/(run)/$|^/(edit)/$|^(/edit(/save|/add|/delete)?)/$")

func viewHandler(w http.ResponseWriter, r *http.Request) {
	p, err := controller.LoadViewPage() //Загрузка данных для генерации страницы
	if err != nil {
		//http.Redirect(w, r, "/edit/"+title, http.StatusFound)
		return
	}
	controller.RenderTemplateView(w, p) //Генерация страницы
}

func resultHandler(w http.ResponseWriter, r *http.Request) {
	a := `https://docs.google.com/spreadsheets/d/1Z7KluzcrGpLu_TVyGUBicwt_mpc1X3FhGlfGluUXh-8/edit` // + "/" + sheetID
	controller.RenderTemplateResult(w, a)                                                           //Генерация страницы
}

func runHandler(w http.ResponseWriter, r *http.Request) {
	switch r.Method {
	case "GET":
		// fmt.Fprintln(w, "Метод GET, копай дальше")
		http.NotFound(w, r)
	case "POST":
		controller.RunProcess(w, r)
		//fmt.Fprintln(w, r.Form)
	}
}

func getFileFormatHandler(w http.ResponseWriter, r *http.Request) {
	switch r.Method {
	case "GET":
		// fmt.Fprintln(w, "Метод GET, копай дальше")
		http.NotFound(w, r)
	case "POST":
		controller.UploadFiles(w, r)
		//fmt.Fprintln(w, r.Form)
	}
}

func makeHandler(fn func(http.ResponseWriter, *http.Request)) http.HandlerFunc {
	return func(w http.ResponseWriter, r *http.Request) {
		m := validPath.FindStringSubmatch(r.URL.Path)
		if m == nil {
			http.NotFound(w, r)
			return
		}
		fn(w, r)
	}
}
