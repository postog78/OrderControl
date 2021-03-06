package excelconnector

import (
	model "OrderControl/app/model"
	"path"
	"time"
)
//ConnectorExcelReader тип, объединяющий общие методы, которые должны быть у коннекторов к данным
type ConnectorExcelReader interface {
	Read(start time.Time, end time.Time) ([]model.TypeReport, error)
	Init()
	GetData() []model.TypeReport
	GetName() string
}

var pathToShipments string = path.Join("Файлы", "Отгрузки")
var pathToArhive string = path.Join("Файлы", "Архивы")
