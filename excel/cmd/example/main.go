package main

import (
	"log"

	"github.com/nonsugar-go/tools/excel"
)

func main() {
	e, err := excel.New("output.xlsx", 11)
	if err != nil {
		log.Fatal(err)
	}
	defer func() {
		if err := e.SaveAndClose(); err != nil {
			log.Print(err)
		}
	}()
	if err := e.NewSheet(
		"気温の一覧", excel.SheetTypeCover); err != nil {
		log.Print(err)
	}
	if err := e.NewSheet(
		"WB シティの気温", excel.SheetTypeNormal); err != nil {
		log.Print(err)
	}
	if err := e.LF().SetVal("WB シティの気温"); err != nil {
		log.Print(err)
	}
	if err := e.SetStyle(excel.NewStyle().Bold()); err != nil {
		log.Print(err)
	}
	rows := [][]any{
		{"月", "気温"},
		{1, 3.5},
		{2, 4.2},
		{3, 6.7},
		{4, 12.8},
		{5, 14.9},
	}
	e.CR(2).LF()
	for _, row := range rows {
		if err := e.LF().SetRow(&row); err != nil {
			log.Print(err)
		}
	}
	e.NewSheet("目次", excel.SheetTypeTOC)
}
