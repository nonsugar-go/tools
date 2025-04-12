package main

import (
	"fmt"
	"log"

	"github.com/nonsugar-go/tools/tui"
)

func main() {
	tui.Title("TUI のデモ")
	tui.MsgBox("TUI の各種関数を一覧する")
	tui.PressAnyKey()
	items := []string{"PaloAlto", "FortiGate"}
	selected, err := tui.Select("機種を選択してください", items)
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println(selected)
	tui.PressAnyKey("Press Any key to continue...")
}
