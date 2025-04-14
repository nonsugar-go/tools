package excel

import (
	"fmt"
	"path/filepath"
	"testing"
)

func TestExcel_NewExcel(t *testing.T) {
	dir := t.TempDir()
	filename := filepath.Join(dir, "Book1.xlsx")
	e, err := NewExcel(filename)
	if err != nil {
		t.Fatalf("NewExcel: want no error, but %v", err)
	}
	defer func() {
		if err := e.Close(); err != nil {
			t.Errorf("Close: want no error, but %v", err)
		}
	}()
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
	t.Logf("Default Font: %s", defFont)
}

func TestExcel_SaveAndClose(t *testing.T) {
	dir := "./testdata"
	filename := filepath.Join(dir, "output.xlsx")
	e, err := NewExcel(filename)
	if err != nil {
		t.Fatalf("NewExcel: want no error, but %v", err)
	}
	defer func() {
		if err := e.SaveAndClose(); err != nil {
			t.Errorf("SaveAndClose: want no error, but %v", err)
		}
	}()
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
		if err := e.SetRow(&row); err != nil {
			t.Errorf("SetRow: want no error, but %v", err)
		}
	}
	if err := e.AddTable("Table1"); err != nil {
		t.Errorf("AddTable: want no error, but %v", err)
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
		got, err := CordinatesToCellName(tt.col, tt.row, tt.abs)
		if !tt.isErr && err != nil {
			t.Errorf("%s: want no error, but %v", tt.name, err)
		}
		if got != tt.expect {
			t.Errorf("%s: want %s, but %s", tt.name, tt.expect, got)
		}
	}
}
