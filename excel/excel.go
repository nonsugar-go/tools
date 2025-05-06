//go:generate stringer -type SheetType -trimprefix SheetType
package excel

import (
	"fmt"
	"sort"

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

	// Cell Sytle
	cellStyleIDs map[cellStyle]int
	cellStyleMap map[string]cellStyle // cellStyleMap["A1"] = CellStyleBold
}

// New returns a pointer to Excel.
func New(book string) (*Excel, error) {
	e := &Excel{
		f:            excelize.NewFile(),
		book:         book,
		Col:          1,
		Row:          1,
		cellStyleIDs: make(map[cellStyle]int),
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
	if e.sheet != "" {
		// 直前のシートに対する処理
		if err := e.applyCellStyle(); err != nil {
			return fmt.Errorf("operation failed on the previous sheet: %s: %w",
				e.sheet, err)
		}
	}
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
	isFoundDefaultSheet := false
	if e.sheet == "" {
		isFoundDefaultSheet = true
	} else {
		// 直前のシートに対する処理
		if err := e.applyCellStyle(); err != nil {
			return fmt.Errorf("operation failed on the previous sheet: %s: %w",
				e.sheet, err)
		}
	}
	_, err := e.f.NewSheet(sheet)
	if err != nil {
		return fmt.Errorf("cannot add the sheet: %s: %w", sheet, err)
	}
	e.sheet = sheet
	e.Col, e.Row = 1, 1
	e.cellStyleMap = make(map[string]cellStyle)
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
	if isFoundDefaultSheet {
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

// Cell obtains the position of a cell in A1 reference format.
func (e *Excel) Cell() (string, error) {
	cell, err := excelize.CoordinatesToCellName(e.Col, e.Row)
	if err != nil {
		return "", fmt.Errorf(
			"failed to convert coordinates to cell name: sheet=%s, col=%d, row=%d: %w",
			e.sheet, e.Col, e.Row, err)
	}
	return cell, nil
}

// CR resets the column to a default value or a specified value.
//
// Example:
//
//	e.CR()  // Resets e.Col to 1.
//	e.CR(2) // Sets e.Col to 2.
func (e *Excel) CR(col ...int) *Excel {
	e.Col = 1
	if len(col) > 0 {
		e.Col = col[0]
	}
	return e
}

// LF increments the row by 1 or a specified amount.
//
// Example:
//
//	e.LF()          // Increments e.Row by 1.
//	e.LF(3)         // Increments e.Row by 3.
//	e.CR().LF(2)    // Resets column and increments row by 2.
func (e *Excel) LF(add ...int) *Excel {
	inc := 1
	if len(add) > 0 {
		inc = add[0]
	}
	e.Row += inc
	return e
}

// SetVal sets a value to a specified cell with optional column
// and row adjustment.
//
// Example:
//
//	err := e.SetVal("Hello, World!", col, row)
func (e *Excel) SetVal(value any, colRow ...int) error {
	if len(colRow) > 0 {
		e.Col = colRow[0]
		if len(colRow) > 1 {
			e.Row = colRow[1]
		}
	}
	cell, err := e.Cell()
	if err != nil {
		return err
	}
	if err := e.f.SetCellValue(e.sheet, cell, value); err != nil {
		return fmt.Errorf("failed to set the cell value: %w", err)
	}
	return nil
}

// SetRow sets row.
//
// Example:
//
//	err := e.SetRow(&[]any{"1", nil, 2})
func (e *Excel) SetRow(row any) error {
	cell, err := e.Cell()
	if err != nil {
		return err
	}
	if err := e.f.SetSheetRow(e.sheet, cell, row); err != nil {
		return fmt.Errorf("failed to set the row data: %w", err)
	}
	return nil
}

// GetLastColumnNumber returns the last column number in the specified sheet.
func (e *Excel) GetLastColumnNumber(sheet ...string) (int, error) {
	sheetName := e.sheet
	if len(sheet) > 0 {
		sheetName = sheet[0]
	}
	rows, err := e.f.GetRows(sheetName)
	if err != nil {
		return 0, fmt.Errorf("failed to retrieve rows from sheet '%s': %w",
			sheetName, err)
	}
	lastCols := 0
	for _, cols := range rows {
		numOfCols := len(cols)
		if lastCols < numOfCols {
			lastCols = numOfCols
		}
	}
	return lastCols, nil
}

// GetLastRowNumber returns the last row number in the specified sheet.
func (e *Excel) GetLastRowNumber(sheet ...string) (int, error) {
	sheetName := e.sheet
	if len(sheet) > 0 {
		sheetName = sheet[0]
	}
	rows, err := e.f.GetRows(sheetName)
	if err != nil {
		return 0, fmt.Errorf("failed to retrieve rows from sheet '%s': %w",
			sheetName, err)
	}
	return len(rows), nil
}

// GetSortedComments retrieves all comments from a sheet and sorts them
// by row and column.
func (e *Excel) GetSortedComments(sheet string) ([]excelize.Comment, error) {
	// Get all comments from the sheet
	comments, err := e.f.GetComments(sheet)
	if err != nil {
		return nil, fmt.Errorf("failed to get comments from sheet '%s': %w",
			sheet, err)
	}

	// Sort comments by row and column
	sort.Slice(comments, func(i, j int) bool {
		colI, rowI, _ := excelize.SplitCellName(comments[i].Cell)
		colJ, rowJ, _ := excelize.SplitCellName(comments[j].Cell)
		if rowI != rowJ {
			return rowI < rowJ
		}
		return colI < colJ
	})

	return comments, nil
}

// AddComment adds a comment to an Excel cell.
func (e *Excel) AddComment(comment string) error {
	cell, err := e.Cell()
	if err != nil {
		return err
	}
	return e.f.AddComment(e.sheet, excelize.Comment{
		Cell:   cell,
		Author: "TOMATO",
		Paragraph: []excelize.RichTextRun{
			{Text: comment, Font: &excelize.Font{
				Bold: false, Italic: false, Underline: "none",
				Family: "MS P ゴシック", Size: 9, Strike: false, Color: "",
				ColorIndexed: 81, ColorTheme: (*int)(nil), ColorTint: 0,
				VertAlign: "",
			}},
		},
	})
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
		e.CR()
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

// CoordinatesToCellName is identical to the excelize module.
func CoordinatesToCellName(col, row int, abs ...bool) (string, error) {
	return excelize.CoordinatesToCellName(col, row, abs...)
}

// ColumnNumberToName is identical to the excelize module.
func ColumnNumberToName(num int) (string, error) {
	return excelize.ColumnNumberToName(num)
}

// RelCellNameToAbsCellName is convert to absolute reference.
func RelCellNameToAbsCellName(cell string) (string, error) {
	col, row, err := excelize.CellNameToCoordinates(cell)
	if err != nil {
		return "", err
	}
	return excelize.CoordinatesToCellName(col, row, true)
}
