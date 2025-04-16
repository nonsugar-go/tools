//go:generate stringer -type SheetType -trimprefix SheetType
package excel

import (
	"fmt"

	"github.com/mattn/go-runewidth"
	"github.com/xuri/excelize/v2"
)

// SheetType indicates the sheet type.
type SheetType int

const (
	SheetTypeUnknown SheetType = iota
	SheetTypeNormal
	SheetTypeTOC
	SheetTypeCover
)

const (
	defaultSheet string = "Sheet1"
	defaultFont  string = "游ゴシック"
)

var (
	codeName         string  = ""
	date1094         bool    = false
	filterPrivacy    bool    = false
	boolTrue         bool    = true
	boolFalse        bool    = false
	int0x0           int     = 0x0
	uint80x0         uint8   = 0x0
	defaultColWidth  float64 = 2.69921875
	defaultRowHeight float64 = 13.5
)

// Excel is a struct that manipulates Excel workbooks.
type Excel struct {
	f         *excelize.File
	book      string
	sheet     string
	sheetType SheetType
	// Current column and Row number
	Col, Row int
}

// NewExcel returns a pointer to Excel.
func NewExcel(book string) (*Excel, error) {
	e := &Excel{
		f:    excelize.NewFile(),
		book: book,
		Col:  1,
		Row:  1,
	}
	if err := e.f.SetDefaultFont(defaultFont); err != nil {
		return nil, err
	}
	if err := e.f.SetWorkbookProps(&excelize.WorkbookPropsOptions{
		CodeName:      &codeName,
		Date1904:      &date1094,
		FilterPrivacy: &filterPrivacy,
	}); err != nil {
		return nil, err
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
		Col:  1,
		Row:  1,
	}
	return e, nil
}

// GetFile returns excelize.File.
func (e *Excel) GetFile() *excelize.File {
	return e.f
}

// Close closes Excel.
func (e *Excel) Close() error {
	if e.f == nil {
		return nil
	}
	if err := e.f.Close(); err != nil {
		return fmt.Errorf(
			"cannot close excel: %s: %w",
			e.book, err)
	}
	e.f = nil
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
func (e *Excel) NewSheet(sheet string, typ ...SheetType) error {
	_, err := e.f.NewSheet(sheet)
	if err != nil {
		return fmt.Errorf(
			"cannot add the sheet: %s: %w",
			sheet, err)
	}
	e.sheet = sheet
	sheetType := SheetTypeUnknown
	if len(typ) > 0 {
		sheetType = typ[0]
	}
	e.sheetType = sheetType
	switch sheetType {
	case SheetTypeNormal:
		fallthrough
	case SheetTypeTOC:
		fallthrough
	case SheetTypeCover:
		if err := e.f.SetSheetProps(
			e.sheet,
			&excelize.SheetPropsOptions{
				AutoPageBreaks:                    &boolTrue,
				BaseColWidth:                      &uint80x0,
				CodeName:                          (*string)(nil),
				CustomHeight:                      &boolTrue,
				DefaultColWidth:                   &defaultColWidth,
				DefaultRowHeight:                  &defaultRowHeight,
				EnableFormatConditionsCalculation: &boolTrue,
				FitToPage:                         (*bool)(nil),
				OutlineSummaryBelow:               &boolTrue,
				OutlineSummaryRight:               (*bool)(nil),
				Published:                         &boolTrue,
				TabColorIndexed:                   (*int)(nil),
				TabColorRGB:                       (*string)(nil),
				TabColorTheme:                     (*int)(nil),
				TabColorTint:                      (*float64)(nil),
				ThickBottom:                       &boolFalse,
				ThickTop:                          &boolFalse,
				ZeroHeight:                        &boolFalse,
			},
			/*
				AutoPageBreak                      true
				BaseColWidth                       0x0
				CodeName                           (*string)(nil)
				CustomHeight                       true
				DefaultColWidth                    2.69921875
				DefaultRowHeight                   13.5
				EnableFormatConditionsCalculation  true
				FitToPage                          (*bool)(nil)
				OutlineSummaryBelow                true
				OutlineSummaryRight                (*bool)(nil)
				Published                          true
				TabColorIndexed                    (*int)(nil)
				TabColorRGB                        (*string)(nil)
				TabColorTheme                      (*int)(nil)
				TabColorTint                       (*float64)(nil)
				ThickBottom                        false
				ThickTop                           false
				ZeroHeight                         false
			*/
		); err != nil {
			return fmt.Errorf("cannot set sheet props: %s: %w",
				sheet, err)
		}
	}
	if e.sheet == "" {
		if err := e.f.DeleteSheet(defaultSheet); err != nil {
			return fmt.Errorf(
				"cannot delete the sheet: %s: %w",
				sheet, err)
		}
	}
	return nil
}

// SetActiveSheet activates the current sheet.
func (e *Excel) SetActiveSheet() error {
	index, err := e.f.GetSheetIndex(e.sheet)
	if err != nil {
		return fmt.Errorf("SetActiveSheet: %w", err)
	}
	e.f.SetActiveSheet(index)
	return nil
}

// SetRow sets row.
//
// Example:
//
//	err := e.SetRow(&[]any{"1", nil, 2})
func (e *Excel) SetRow(row any) error {
	cell, err := excelize.CoordinatesToCellName(e.Col, e.Row)
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
	e.Row++
	return nil
}

// Header is a structure consisting of table column names and column widths.
type Header struct {
	Text  string
	Width float64
}

// SetHeader specifies table column names and column widths.
func (e *Excel) SetHeader(headers []Header) error {
	for col, header := range headers {
		if header.Width <= 0 {
			header.Width = float64(runewidth.StringWidth(header.Text)) + 1.7
		}
		cell, err := excelize.CoordinatesToCellName(col+1, 1)
		if err != nil {
			return fmt.Errorf("SetHeader: %w", err)
		}
		if err := e.f.SetCellStr(e.sheet, cell, header.Text); err != nil {
			return fmt.Errorf("SetHeader: %w", err)
		}
		colName, err := excelize.ColumnNumberToName(col + 1)
		if err != nil {
			return fmt.Errorf("SetHeader: %w", err)
		}
		if err := e.f.SetColWidth(e.sheet, colName, colName, header.Width); err != nil {
			return fmt.Errorf("SetHeader: %w", err)
		}
		e.Col, e.Row = 1, 2
	}
	return nil
}

// AddTable adds an Excel table.
func (e *Excel) AddTable(table string) error {
	rows, err := e.f.GetRows(e.sheet)
	if err != nil {
		return err
	}
	lastRow := len(rows)
	lastCol := 0
	if lastRow > 0 {
		lastCol = len(rows[0])
	}
	cellName, err := excelize.CoordinatesToCellName(lastCol, lastRow)
	if err != nil {
		return fmt.Errorf("AddTable: %w", err)
	}
	err = e.f.AddTable(
		e.sheet,
		&excelize.Table{
			Range:     "A1:" + cellName,
			Name:      table,
			StyleName: "TableStyleLight16",
		},
	)
	if err != nil {
		return fmt.Errorf("AddTable: %w", err)
	}
	return nil
}

// CordinatesToCellName is identical to the excelize module.
func CordinatesToCellName(col, row int, abs ...bool) (string, error) {
	return excelize.CoordinatesToCellName(col, row, abs...)
}

// ColumnNumberToName is identical to the excelize module.
func ColumnNumberToName(num int) (string, error) {
	return excelize.ColumnNumberToName(num)
}
