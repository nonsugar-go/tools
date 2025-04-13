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
