package excelconnector

import (
	"path"
)

//IstraOpt коннектор Истры опта
type IstraOpt struct {
	Istra
}

//Init инициализация коннектора к экселю
func (basis *IstraOpt) Init() {
	basis.Name = "Истра опт"
	basis.PathToDir = path.Join(pathToShipments, "ИстраОпт")
	basis.colComment = 10
	basis.colVolume = 15
	basis.colWeight = 14
	basis.colDate = 1
	basis.colTypeOfProduct = 12
}
