package excel

import (
	"fmt"
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestExcel_NewExcel(t *testing.T) {
	dir := t.TempDir()
	filename := filepath.Join(dir, "Book1.xlsx")
	tests := []struct {
		name     string
		fontSize float64
		ok       bool
	}{
		{"size=11", 11, true},
		{"size=10", 10, true},
		{"size=409", 409, true},
		{"size=10.5", 10.5, true},
		{"size=10.4", 10.4, false},
		{"size=409.5", 409.5, false},
		{"size=0", 0, false},
		{"size=0", -1, false},
	}
	for _, tt := range tests {
		e, err := New(filename, tt.fontSize)
		if tt.ok && err != nil {
			t.Errorf("NewExcel: want no error, but %v", err)
		}
		if !tt.ok && err == nil {
			t.Errorf("NewExcel: want error, but %v", err)
		}
		if e != nil {
			if err := e.Close(); err != nil {
				t.Errorf("Close: want no error, but %v", err)
			}
		}
	}
}

func TestExcel_OpenExcel(t *testing.T) {
	dir := "./testdata"
	filename := filepath.Join(dir, "Sample1.xlsx")
	e, err := OpenExcel(filename)
	if err != nil {
		t.Fatalf("OpenExcel: want no error, but %v", err)
	}
	defer func() {
		if err := e.Close(); err != nil {
			t.Errorf("Close: want no error, but %v", err)
		}
	}()
	file := e.GetFile()
	defFont, _ := file.GetDefaultFont()
	if defFont != defaultFont {
		t.Errorf("GetDefaultfont: want %s, but %s", defaultFont, defFont)
	}
}

func TestExcel_SaveAndClose(t *testing.T) {
	dir := "./testdata"
	filename := filepath.Join(dir, "output.xlsx")
	e, err := New(filename)
	if err != nil {
		t.Fatalf("NewExcel: want no error, but %v", err)
	}
	defer func() {
		if err := e.SaveAndClose(); err != nil {
			t.Errorf("SaveAndClose: want no error, but %v", err)
		}
	}()

	// シート「サンプル テスト」
	if err := e.NewSheet("サンプル テスト"); err != nil {
		t.Errorf("NewSheet: want no error, but: %v", err)
	}
	tests := []struct {
		name string
		x, y int
	}{
		{name: "(13, 3)", x: 13, y: 3},
		{name: "(14, 4)", x: 14, y: 4},
	}
	for _, tt := range tests {
		e.Col = tt.x
		e.Row = tt.y
		err := e.SetRow(&[]any{
			fmt.Sprintf("(%d, %d)", e.Col, e.Row),
			"1", nil, 2})
		if err != nil {
			t.Errorf("SetRow: %s: want no error, but %v", tt.name, err)
		}
	}

	// シート「人物リスト」
	if err := e.NewSheet("人物リスト"); err != nil {
		t.Errorf("NewSheet: want no error, but: %v", err)
	}
	if err := e.SetActiveSheet(); err != nil {
		t.Errorf("NewActiveSheet: want no error, but: %v", err)
	}
	if err := e.SetHeader([]Header{
		{"No", 0},
		{"姓", 6},
		{"名", 0},
		{"人物を説明することば", 0},
	}); err != nil {
		t.Errorf("SetHader: want no error, but %v", err)
	}
	rows := [][]any{
		{1, "大谷", "翔平", "ベースボール プレイヤー"},
		{2, "鈴木", "一郎", "殿堂入り"},
		{3, "野茂", "英雄", "トルネード投法"},
	}
	for _, row := range rows {
		if err := e.LF().SetRow(&row); err != nil {
			t.Errorf("SetRow: want no error, but %v", err)
		}
	}
	if err := e.AddTable("Table1"); err != nil {
		t.Errorf("AddTable: want no error, but %v", err)
	}

	// シート「設定シート」
	if err := e.NewSheet("設定シート", SheetTypeTOC); err != nil {
		t.Errorf("NewSheet: want no error, but: %v", err)
	}
	if err := e.SetVal("設定シート"); err != nil {
		t.Errorf("SetVal: want no error, but: %v", err)
	}
	if err := e.MarkHeader(1); err != nil {
		t.Errorf("MarkHeader: want no error, but: %v", err)
	}
	if err := e.SetVal("セル スタイル", 1, 5); err != nil {
		t.Errorf("SetVal: want no error, but: %v", err)
	}
	if err := e.MarkHeader(2); err != nil {
		t.Errorf("MarkHeader: want no error, but: %v", err)
	}
	e.CR(2).LF(2)
	if err := e.SetRow(&[]any{
		"cellStyleIDs", nil, nil, nil, nil, e.cellStyleIDs}); err != nil {
		t.Errorf("SetRow: want no error, but: %v", err)
	}
	if err := e.LF().SetRow(&[]any{
		"cellStyleMap", nil, nil, nil, nil, e.cellStyleMap}); err != nil {
		t.Errorf("SetRow: want no error, but: %v", err)
	}

	// シート「標準」
	if err := e.NewSheet("標準", SheetTypeNormal); err != nil {
		t.Errorf("NewSheet: want no error, but: %v", err)
	}
	if err := e.SetVal("標準"); err != nil {
		t.Errorf("SetVal: want no error, but: %v", err)
	}
	if err := e.MarkHeader(1); err != nil {
		t.Errorf("MarkHeader: want no error, but: %v", err)
	}
	if err := e.SetVal("レベル2のヘッダ", 1, 5); err != nil {
		t.Errorf("SetVal: want no error, but: %v", err)
	}
	if err := e.MarkHeader(2); err != nil {
		t.Errorf("MarkHeader: want no error, but: %v", err)
	}
	e.CR(2).LF(2)
	if err := e.SetRow(&[]any{"数", 12, 13, 14, 15, 16}); err != nil {
		t.Errorf("SetRow: want no error, but: %v", err)
	}
	if err := e.LF().SetRow(&[]any{"このセルには文章を入力します。"}); err != nil {
		t.Errorf("SetRow: want no error, but: %v", err)
	}
	if err := e.SetCellStyleForCurrentCell(
		NewStyle().Bold().Add(bbT, bbL, bB),
	); err != nil {
		t.Errorf("SetCellStyleForCurrentCell: want no error, but: %v", err)
	}
	for y := 10; y < 20; y++ {
		for x := 3; x < 16; x++ {
			cell, err := excelize.CoordinatesToCellName(x, y)
			if err != nil {
				t.Errorf("want no error, but: %v", err)
			}
			if y == 10 {
				if err = e.SetCellStyle(cell, NewStyle().Add(bbT)); err != nil {
					t.Errorf("want no error, but: %v", err)
				}
			}
			if y == 19 {
				if err = e.SetCellStyle(cell, NewStyle().Add(bbB)); err != nil {
					t.Errorf("want no error, but: %v", err)
				}
			}
			if x == 3 {
				if err = e.SetCellStyle(cell, NewStyle().Add(bbL)); err != nil {
					t.Errorf("want no error, but: %v", err)
				}
			}
			if x == 15 {
				if err = e.SetCellStyle(cell, NewStyle().Add(bbR)); err != nil {
					t.Errorf("want no error, but: %v", err)
				}
			}
		}
	}
	if err := e.AddComment("コメントの例です"); err != nil {
		t.Errorf("AddComment: want no error, but: %v", err)
	}
	if err := e.SetVal("レベル3のヘッダ", 1, 11); err != nil {
		t.Errorf("SetVal: want no error, but: %v", err)
	}
	if err := e.MarkHeader(3); err != nil {
		t.Errorf("MarkHeader: want no error, but: %v", err)
	}
	e.LF(2)
	if err := e.SetVal("定数 CellStyle"); err != nil {
		t.Errorf("SetVal: want no error, but: %v", err)
	}
	if err := e.MarkHeader(2); err != nil {
		t.Errorf("MarkHeader: want no error, but: %v", err)
	}
	e.CR(2).LF()
	for _, v := range []any{
		cellStyleNormal,
		cellStyleBold,
		bT, bL, bR, bB,
		bbT, bbL, bbR, bbB,
	} {
		if err := e.LF().SetRow(&[]any{v}); err != nil {
			t.Errorf("SetRow: want no error, but: %v", err)
		}
	}
	e.CR().LF(3)
	if err := e.SetVal("CellStyle map"); err != nil {
		t.Errorf("SetVal: want no error, but: %v", err)
	}
	if err := e.MarkHeader(3); err != nil {
		t.Errorf("MarkHeader: want no error, but: %v", err)
	}
	e.CR(2).LF()
	if err := e.LF().SetRow(&[]any{
		"cellStyleIDs", nil, nil, nil, nil, e.cellStyleIDs}); err != nil {
		t.Errorf("SetRow: want no error, but: %v", err)
	}
	if err := e.LF().SetRow(&[]any{
		"cellStyleMap", nil, nil, nil, nil, e.cellStyleMap}); err != nil {
		t.Errorf("SetRow: want no error, but: %v", err)
	}
	e.CR().LF(3)
	if err := e.SetVal("背景の例"); err != nil {
		t.Errorf("SetVal: want no error, but: %v", err)
	}
	if err := e.MarkHeader(1); err != nil {
		t.Errorf("MarkHeader: want no error, but: %v", err)
	}
	e.CR(2).LF()
	for _, bg := range []struct {
		name  string
		style cellStyle
	}{
		{"灰色1", fillGray1},
		{"灰色2", fillGray2},
		{"灰色3", fillGray3},
		{"ピンク", fillPink},
		{"黄色", fillYellow},
		{"薄い青", fillLightBlue},
	} {
		if err := e.LF().SetVal(bg.name); err != nil {
			t.Errorf("SetVal: want no error, but: %v", err)
		}
		if err := e.SetCellStyleForCurrentCell(
			NewStyle().Add(bg.style),
		); err != nil {
			t.Errorf("SetCellStyleForCurrentCell: want no error, but: %v", err)
		}
	}
	e.CR().LF(3)
	if err := e.SetVal("コメントの一覧"); err != nil {
		t.Errorf("SetVal: want no error, but: %v", err)
	}
	if err := e.MarkHeader(1); err != nil {
		t.Errorf("MarkHeader: want no error, but: %v", err)
	}
	e.CR(2).LF()
	comments, err := e.GetSortedComments(e.sheet)
	if err != nil {
		t.Errorf("GetSortedComments: want no error, but: %v", err)
	}
	for _, comment := range comments {
		if err := e.LF().SetRow(&[]any{
			comment.Cell, nil, nil, nil, comment.Paragraph,
		}); err != nil {
			t.Errorf("SetVal: want no error, but: %v", err)
		}
	}

	// シート「目次」
	if err := e.NewSheet("目次", SheetTypeTOC); err != nil {
		t.Errorf("NewSheet: want no error, but: %v", err)
	}
	if err := e.SetVal("目次"); err != nil {
		t.Errorf("SetVal: want no error, but: %v", err)
	}
	if err := e.MarkHeader(1); err != nil {
		t.Errorf("MarkHeader: want no error, but: %v", err)
	}
	e.Col, e.Row = 10, 5
	if err := e.MakeTOC(); err != nil {
		t.Errorf("MarkTOC: want no error, but: %v", err)
	}
}

func Test_ColumnNumberToName(t *testing.T) {
	tests := []struct {
		name   string
		colNum int
		expect string
		isErr  bool
	}{
		{"col number is 5", 5, "E", false},
		{"col number is 6", 6, "F", false},
		{"col number is 0", 0, "", true},
		{"col number is -1", -1, "", true},
	}

	for _, tt := range tests {
		got, err := ColumnNumberToName(tt.colNum)
		if !tt.isErr && err != nil {
			t.Errorf("%s: want no error, but %v", tt.name, err)
		}
		if got != tt.expect {
			t.Errorf("%s: want %s, but %s", tt.name, tt.expect, got)
		}
	}
}

func Test_CordinatesToCellName(t *testing.T) {
	tests := []struct {
		name     string
		col, row int
		abs      bool
		expect   string
		isErr    bool
	}{
		{"col: 1, row: 2", 1, 2, false, "A2", false},
		{"col: 3, row: 1", 3, 1, true, "$C$1", false},
		{"col: 0, row: 1", 0, 1, false, "", true},
	}

	for _, tt := range tests {
		got, err := CoordinatesToCellName(tt.col, tt.row, tt.abs)
		if !tt.isErr && err != nil {
			t.Errorf("%s: want no error, but %v", tt.name, err)
		}
		if got != tt.expect {
			t.Errorf("%s: want %s, but %s", tt.name, tt.expect, got)
		}
	}
}

func TestExcelCell(t *testing.T) {
	tests := []struct {
		name     string
		col, row int
		expect   string
		isErr    bool
	}{
		{"col: 1, row: 2", 1, 2, "A2", false},
		{"col: 3, row: 1", 3, 1, "C1", false},
		{"col: 0, row: 1", 0, 1, "", true},
		{"col: 10, row: 99999999999999", 0, 1, "", true},
	}

	e, _ := New("dummy.xlsx")
	defer e.Close()
	for _, tt := range tests {
		e.Col, e.Row = tt.col, tt.row
		got, err := e.Cell()
		if !tt.isErr && err != nil {
			t.Errorf("%s: want no error, but %v", tt.name, err)
		}
		if got != tt.expect {
			t.Errorf("%s: want %s, but %s", tt.name, tt.expect, got)
		}
	}
}

func TestExcelMakeHeader(t *testing.T) {
	tests := []struct {
		name  string
		level int
		isErr bool
	}{
		{"level=-1", -1, true},
		{"level=0", 0, false},
		{"level=1", 1, false},
		{"level=3", 3, false},
		{"level=4", 4, true},
	}

	e, _ := New("dummy.xlsx")
	defer e.Close()
	_ = e.NewSheet("foo")
	for _, tt := range tests {
		err := e.MarkHeader(tt.level)
		if !tt.isErr && err != nil {
			t.Errorf("%s: want no error, but %v", tt.name, err)
		}
		if tt.isErr && err == nil {
			t.Errorf("%s: want error, but %v", tt.name, err)
		}
	}
}

func TestExcelGetLastColumnNumberAndGetLastRowNumber(t *testing.T) {
	tests := []struct {
		name     string
		sheet    string
		col, row int
	}{
		{"test1 1 5", "test 1", 1, 5},
		{"test2 3 8", "test 2", 3, 8},
		{"test3 9 12", "test 3", 9, 12},
	}

	e, _ := New("dummy.xlsx")
	defer e.Close()
	for _, tt := range tests {
		_ = e.NewSheet(tt.name)
		_ = e.SetVal(tt.name, tt.col, tt.row)
		lastC, err := e.GetLastColumnNumber()
		if err != nil {
			t.Errorf("%s: GetLastColumnNumber: want no error, but %v",
				tt.name, err)
		}
		lastR, err := e.GetLastRowNumber()
		if err != nil {
			t.Errorf("%s: GetLastRowNumber: want no error, but %v", tt.name, err)
		}
		if lastC != tt.col {
			t.Errorf("%s: GetLastColumnNumber: want %d, but %d",
				tt.name, tt.col, lastC)
		}
		if lastR != tt.row {
			t.Errorf("%s: GetLastColumnNumber: want %d, but %d",
				tt.name, tt.row, lastR)
		}
	}

	tt := tests[0]
	lastC, err := e.GetLastColumnNumber(tt.name)
	if err != nil {
		t.Errorf("%s: GetLastColumnNumber: want no error, but %v",
			tt.name, err)
	}
	lastR, err := e.GetLastRowNumber(tt.name)
	if err != nil {
		t.Errorf("%s: GetLastRowNumber: want no error, but %v", tt.name, err)
	}
	if lastC != tt.col {
		t.Errorf("%s: GetLastColumnNumber: want %d, but %d",
			tt.name, tt.col, lastC)
	}
	if lastR != tt.row {
		t.Errorf("%s: GetLastColumnNumber: want %d, but %d",
			tt.name, tt.row, lastR)
	}
}

func TestRelCellNameToAbsCellName(t *testing.T) {
	tests := []struct {
		name     string
		cell     string
		expected string
	}{
		{"A1", "A1", "$A$1"},
		{"BZ1234", "BZ1234", "$BZ$1234"},
		{"$H$20", "$H$20", "$H$20"},
	}

	e, _ := New("dummy.xlsx")
	defer e.Close()
	_ = e.NewSheet("dummy")
	for _, tt := range tests {
		got, err := RelCellNameToAbsCellName(tt.cell)
		if err != nil {
			t.Errorf("%s: want no error, but %v",
				tt.name, err)
		}
		if got != tt.expected {
			t.Errorf("%s: want %s, but %s",
				tt.name, tt.expected, got)
		}
	}
}
