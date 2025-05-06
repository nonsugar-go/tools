package main

import (
	"log"

	"github.com/nonsugar-go/tools/excel"
)

func main() {
	e, err := excel.New("output.xlsx")
	if err != nil {
		log.Fatal(err)
	}
	defer func() {
		if err := e.SaveAndClose(); err != nil {
			log.Print(err)
		}
	}()
	if err := e.NewSheet("WB シティの気温"); err != nil {
		log.Print(err)
	}
	if err := e.SetVal("WB シティの気温"); err != nil {
		log.Print(err)
	}
	if err := e.SetCellStyle(excel.NewStyle().Bold()); err != nil {
		log.Print(err)
	}
	rows := [][]any{
		{"WB シティの気温"},
		{},
		{"月", "気温"},
		{1, 3.5},
		{2, 4.2},
		{3, 6.7},
		{4, 12.8},
		{5, 14.9},
	}
	e.LF()
	for _, row := range rows {
		if err := e.LF().SetRow(&row); err != nil {
			log.Print(err)
		}
	}
}
