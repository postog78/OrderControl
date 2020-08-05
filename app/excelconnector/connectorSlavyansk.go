package excelconnector

import (
	"path"
)

//MainSpecialConstruction коннектор Главспецстроя
type Slavyansk struct {
	FlatTable
}

//Init инициализация коннектора к экселю
func (basis *Slavyansk) Init() {
	basis.Name = "Славянск"
	basis.PathToDir = path.Join(pathToShipments, "Славянск")

	basis.rowBegin = 3
	basis.colDate = 1
	basis.colComment = 4
	basis.colVolume = 6
	basis.colWeight = 7
	basis.weightInTone = false
	basis.colTypeOfProduct = 7
	basis.colDriver = 5
}
