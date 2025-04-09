package excel

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

const (
	defaultSheet = "Sheet1"
)

type Excel struct {
	f        *excelize.File
	book     string
	sheet    string
	col, row int
}

func NewExcel(book string) (*Excel, error) {
	e := &Excel{
		f:    excelize.NewFile(),
		book: book,
		col:  1,
		row:  1,
	}
	return e, nil
}

func (e *Excel) Close() error {
	if err := e.f.Close(); err != nil {
		return fmt.Errorf(
			"cannot excelize.File: %s: %s",
			e.book, err.Error())
	}
	return nil
}

func (e *Excel) SaveAndClose() error {
	if err := e.f.SaveAs(e.book); err != nil {
		return fmt.Errorf(
			"cannot save excel book: %s: %s",
			e.book, err.Error())
	}
	if err := e.Close(); err != nil {
		return err
	}
	return nil
}

func (e *Excel) NewSheet(sheet string) error {
	_, err := e.f.NewSheet(sheet)
	if err != nil {
		return fmt.Errorf(
			"cannot add the sheet: %s: %s",
			sheet, err.Error())
	}
	if e.sheet == "" {
		if err := e.f.DeleteSheet(defaultSheet); err != nil {
			return fmt.Errorf(
				"cannot delete the sheet: %s: %s",
				sheet, err.Error())
		}
	}
	e.sheet = sheet
	return nil
}

func (e *Excel) SetRow(row any) error {
	cell, err := excelize.CoordinatesToCellName(e.col, e.row)
	if err != nil {
		return fmt.Errorf(
			"cannot coordinates to cell name: %s: %s",
			e.sheet, err.Error())
	}
	err = e.f.SetSheetRow(e.sheet, cell, row)
	if err != nil {
		return fmt.Errorf(
			"cannot set the row: %s: %s",
			e.sheet, err.Error())
	}
	e.row++
	return nil
}
