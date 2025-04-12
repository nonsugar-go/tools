package main

import (
	"fmt"
	"log"

	"github.com/nonsugar-go/tools/tui"
)

func main() {
	selected, err := tui.FilePicker(
		"go のソースファイルを選択してください",
		[]string{".go"})
	if err != nil {
		log.Print("ファイルが選択されませんでした")
	} else {
		fmt.Println("選択されたファイル:", selected)
	}
}
