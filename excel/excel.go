package excel

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

const (
	defaultSheet = "Sheet1"
)

// Excel is a struct that manipulates Excel workbooks.
type Excel struct {
	f        *excelize.File
	book     string
	sheet    string
	col, row int
}

// NewExcel returns a pointer to Excel.
func NewExcel(book string) (*Excel, error) {
	e := &Excel{
		f:    excelize.NewFile(),
		book: book,
		col:  1,
		row:  1,
	}
	return e, nil
}

// OpenExcel opens Excel workbooks
func OpenExcel(book string) (*Excel, error) {
	f, err := excelize.OpenFile(book)
	if err != nil {
		return nil, fmt.Errorf(
			"cannot open file: %s, %w",
			book, err)
	}
	e := &Excel{
		f:    f,
		book: book,
		col:  1,
		row:  1,
	}
	return e, nil
}

// Close closes Excel.
func (e *Excel) Close() error {
	if err := e.f.Close(); err != nil {
		return fmt.Errorf(
			"cannot close excel: %s: %w",
			e.book, err)
	}
	return nil
}

// SaveAndClose saves and closes the workbook.
func (e *Excel) SaveAndClose() error {
	if err := e.f.SaveAs(e.book); err != nil {
		return fmt.Errorf(
			"cannot save excel book: %s: %w",
			e.book, err)
	}
	if err := e.Close(); err != nil {
		return err
	}
	return nil
}

// NewSheet creates a new worksheet.
func (e *Excel) NewSheet(sheet string) error {
	_, err := e.f.NewSheet(sheet)
	if err != nil {
		return fmt.Errorf(
			"cannot add the sheet: %s: %w",
			sheet, err)
	}
	if e.sheet == "" {
		if err := e.f.DeleteSheet(defaultSheet); err != nil {
			return fmt.Errorf(
				"cannot delete the sheet: %s: %w",
				sheet, err)
		}
	}
	e.sheet = sheet
	return nil
}

// SetRow sets row.
func (e *Excel) SetRow(row any) error {
	cell, err := excelize.CoordinatesToCellName(e.col, e.row)
	if err != nil {
		return fmt.Errorf(
			"cannot coordinates to cell name: %s: %w",
			e.sheet, err)
	}
	err = e.f.SetSheetRow(e.sheet, cell, row)
	if err != nil {
		return fmt.Errorf(
			"cannot set the row: %s: %w",
			e.sheet, err)
	}
	e.row++
	return nil
}
