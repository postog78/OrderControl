package excelconnector

import (
	"path"
)

//MainSpecialConstruction коннектор Главспецстроя
type MainSpecialConstruction struct {
	FlatTable
}

//Init инициализация коннектора к экселю
func (basis *MainSpecialConstruction) Init() {
	basis.Name = "ГлавСпецСтрой"
	basis.PathToDir = path.Join(pathToShipments, "ГлавСпецСтрой")

	basis.colDate = 0
	basis.colComment = 1
	basis.colVolume = 10
	basis.colWeight = 9
	basis.weightInTone = true
	basis.colTypeOfProduct = 7
	basis.colDriver = 5
}
