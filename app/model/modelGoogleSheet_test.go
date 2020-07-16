package model

import (
	"reflect"
	"testing"
	"time"
)

func Test_getTimeFromTitleSheet(t *testing.T) {
	type args struct {
		title string
	}
	tests := []struct {
		name        string
		args        args
		wantRetTime time.Time
		wantErr     bool
	}{
		{"Норма 1. Продажа Декабрь 2019 (РЕГИОНЫ)", args{"Продажа Декабрь 2019 (РЕГИОНЫ)"}, time.Date(2019, time.December, 1, 0, 0, 0, 0, time.UTC), false},
		{"Норма 2, Продажа Декабрь 2019 (остальные)", args{"Продажа Декабрь 2019 (остальные)"}, time.Date(2019, time.December, 1, 0, 0, 0, 0, time.UTC), false},
		{"Норма 3, Продажа Январь 2020 (РЕГИОНЫ)", args{"Продажа Январь 2020 (РЕГИОНЫ)"}, time.Date(2020, time.January, 1, 0, 0, 0, 0, time.UTC), false},
		{"Норма 4, Продажа МАРТ 2020 (РЕГИОНЫ)", args{"Продажа МАРТ 2020 (РЕГИОНЫ)"}, time.Date(2020, time.March, 1, 0, 0, 0, 0, time.UTC), false},
		// "Норма 5",
		{"Не Норма 1, Загрузка_марийка", args{"Загрузка_марийка"}, time.Time{}, true},
		// "Не норма 2",
		{"Не Норма 2 Пустое значение", args{""}, time.Time{}, true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			gotRetTime, err := getTimeFromTitleSheet(tt.args.title)
			if (err != nil) != tt.wantErr {
				t.Errorf("getTimeFromTitleSheet() error = %v, wantErr %v", err, tt.wantErr)
				return
			}
			if !reflect.DeepEqual(gotRetTime, tt.wantRetTime) {
				t.Errorf("getTimeFromTitleSheet() = %v, want %v", gotRetTime, tt.wantRetTime)
			}
		})
	}
}

// func Test_getMonthBefore(t *testing.T) {
// 	type args struct {
// 		t time.Time
// 	}
// 	tests := []struct {
// 		name string
// 		args args
// 		want time.Time
// 	}{
// 		{"Норма 1. 01/01/2019", args{time.Date(2019, time.January, 1, 0, 0, 0, 0, time.UTC)}, time.Date(2018, time.December, 1, 0, 0, 0, 0, time.UTC)},
// 		{"Норма 2. 15/01/2019", args{time.Date(2019, time.January, 15, 0, 0, 0, 0, time.UTC)}, time.Date(2018, time.December, 1, 0, 0, 0, 0, time.UTC)},
// 		{"Норма 1. 07/06/2019", args{time.Date(2019, time.June, 07, 0, 0, 0, 0, time.UTC)}, time.Date(2019, time.May, 1, 0, 0, 0, 0, time.UTC)},
// 	}
// 	for _, tt := range tests {
// 		t.Run(tt.name, func(t *testing.T) {
// 			if got := getMonthBefore(tt.args.t); !reflect.DeepEqual(got, tt.want) {
// 				t.Errorf("getMonthBefore() = %v, want %v", got, tt.want)
// 			}
// 		})
// 	}
// }
