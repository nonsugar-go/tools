package tui

import (
	"strings"
	"testing"
)

func TestFilePicker(t *testing.T) {
	got, err := FilePicker(
		"go のファイルを選択してください", []string{".go"})
	if len(got) > 0 && err != nil {
		t.Errorf("want no selected file, but got %s, and error: %v", got, err)
	}
	if len(got) != 0 && !strings.HasSuffix(got, ".go") {
		t.Errorf("want *.go, but %s", got)
	}
}
