package excel

import (
	"os"
	"path/filepath"
	"testing"
)

func TestExcel_NewExcel(t *testing.T) {
	dir := t.TempDir()
	filename := filepath.Join(dir, "Book1.xlsx")
	e, err := NewExcel(filename)
	if err != nil {
		t.Fatalf("want no error, but %v", err)
	}
	if err := e.Close(); err != nil {
		t.Errorf("want no error, but %v", err)
	}
}

func TestExcel_SaveAndClose(t *testing.T) {
	// dir := t.TempDir()
	dir := "./testdata"
	filename := filepath.Join(dir, "output.xlsx")
	e, err := NewExcel(filename)
	if err != nil {
		t.Fatalf("want no error, but %v", err)
	}
	if err := e.NewSheet("サンプル テスト"); err != nil {
		t.Errorf("want no error, but: %v", err)
	}
	rows := [][]any{
		{"No", "姓", "名", "備考"},
		{1, "大谷", "翔平", "ベースボール プレイヤー"},
		{2, "鈴木", "一郎", "殿堂入り"},
		{3, "野茂", "英雄", "トルネード投法"},
	}
	for _, row := range rows {
		if err := e.SetRow(&row); err != nil {
			t.Errorf("want no error, but %v", err)
		}
	}
	if err := e.SaveAndClose(); err != nil {
		t.Errorf("want no error, but %v", err)
	}
	fileInfo, err := os.Stat(filename)
	if err != nil {
		t.Errorf("cannot stat: %s: %v",
			filename, err)
	}
	if fileInfo.Size() == 0 {
		t.Fatal("filesize: want >0, but 0")
	}
}

func TestExcel_OpenExcel(t *testing.T) {
	dir := "./testdata"
	filename := filepath.Join(dir, "Sample1.xlsx")
	e, err := OpenExcel(filename)
	if err != nil {
		t.Fatalf("want no error, but %v", err)
	}
	if err := e.Close(); err != nil {
		t.Errorf("want no error, but %v", err)
	}
}
