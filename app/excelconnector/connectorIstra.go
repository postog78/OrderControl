package excelconnector

import (
	"path"
)

//Istra коннектор Истры
type Istra struct {
	FlatTable
}

// var excelFileName string = `C:\Users\Dell\Documents\Go\OrderControl\Files\Отчеты базисов об отгрузках\Базис 1\ИНТ_Остатки 2020.xlsx`

//Init инициализация коннектора к экселю
func (basis *Istra) Init() {
	basis.Name = "Истра"
	basis.PathToDir = path.Join(pathToShipments, "Истра")
	basis.colComment = 9
	basis.colVolume = 13
	basis.colWeight = 12
	basis.colDate = 1
	basis.colTypeOfProduct = 10
	basis.weightInTone = true
	basis.colDriver = 8
}
