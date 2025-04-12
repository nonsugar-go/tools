package main

import (
	"github.com/nonsugar-go/tools/tui"
)

func main() {
	tui.PrintTable(
		[]string{"項目", "設定値"},
		[][]string{
			{"機器の種類", "FortiGate"},
			{"設定ファイル名", "fgt.conf"},
			{"設定表ファイル名", "fgt.xlsx"},
		})
}
