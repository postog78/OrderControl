package excelconnector

import (
	"path"
)

//IstraOpt коннектор Истры опта
type IstraOpt struct {
	FlatTable
}

//Init инициализация коннектора к экселю
func (basis *IstraOpt) Init() {
	basis.Name = "Истра опт"
	basis.PathToDir = path.Join(pathToShipments, "ИстраОпт")
	basis.colComment = 11
	basis.colVolume = 15
	basis.colWeight = 14
	basis.colDate = 1
	basis.colTypeOfProduct = 12
	basis.weightInTone = true
	basis.colDriver = 10
}
