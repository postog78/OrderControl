package model

import "time"

type TypeReport struct {
	numOrder int
	weight   int //кг
	date     time.Time
	volume   int
}
