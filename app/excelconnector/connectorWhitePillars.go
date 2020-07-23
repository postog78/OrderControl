package excelconnector

import (
	"path"
	"time"
)

//WhitePillars коннектор Истры
type WhitePillars struct {
	FlatTable
}

// var excelFileName string = `C:\Users\Dell\Documents\Go\OrderControl\Files\Отчеты базисов об отгрузках\Базис 1\ИНТ_Остатки 2020.xlsx`

//Init инициализация коннектора к экселю
func (basis *WhitePillars) Init() {

	basis.Name = "Белые столбы"
	basis.PathToDir = path.Join(pathToShipments, "Белые столбы")

	basis.rowBegin = 2
	basis.colDate = 0
	basis.colComment = 1
	basis.colVolume = 4
	basis.colWeight = 6
	basis.colTypeOfProduct = 7
	basis.weightInTone = false
}

//getDateFromString Получение даты из строки
func (basis *WhitePillars) getDateFromString(str string) (date time.Time, err error) {
	return time.Parse(`02.01.2006`, str)
}
