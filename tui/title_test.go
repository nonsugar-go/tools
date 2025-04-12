package tui

import (
	"io"
	"os"
	"strings"
	"testing"
)

func TestTitle(t *testing.T) {
	oldOut := os.Stdout
	r, w, _ := os.Pipe()
	os.Stdout = w
	title := "タイトルのテスト"
	Title(title)
	_ = w.Close()
	os.Stdout = oldOut
	out, _ := io.ReadAll(r)
	got := string(out)
	if !strings.Contains(got, title) {
		t.Errorf("want %s, but %s", title, got)
	}
}

func TestMsgBox(t *testing.T) {
	oldOut := os.Stdout
	r, w, _ := os.Pipe()
	os.Stdout = w
	msg := "タイトルのテスト"
	MsgBox(msg)
	_ = w.Close()
	os.Stdout = oldOut
	out, _ := io.ReadAll(r)
	got := string(out)
	if !strings.Contains(got, msg) {
		t.Errorf("want %s, but %s", msg, got)
	}
}
