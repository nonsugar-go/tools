package main

import (
	"fmt"
	"os"

	"github.com/charmbracelet/bubbles/list"
	tea "github.com/charmbracelet/bubbletea"
	"github.com/nonsugar-go/tools/tui"
)

func main() {
	items := []list.Item{
		tui.Item("Ramen"),
		tui.Item("Tomato Soup"),
		tui.Item("Hamburgers"),
		tui.Item("Cheeseburgers"),
		tui.Item("Currywurst"),
		tui.Item("Okonomiyaki"),
		tui.Item("Pasta"),
		tui.Item("Fillet Mignon"),
		tui.Item("Caviar"),
		tui.Item("Just Wine"),
	}

	const defaultWidth = 20

	l := list.New(items, tui.ItemDelegate{}, defaultWidth, tui.ListHeight)
	l.Title = "What do you want for dinner?"
	l.SetShowStatusBar(false)
	l.SetFilteringEnabled(false)
	l.Styles.Title = tui.TitleStyle
	l.Styles.PaginationStyle = tui.PaginationStyle
	l.Styles.HelpStyle = tui.HelpStyle

	m := tui.Model{List: l}

	if _, err := tea.NewProgram(m).Run(); err != nil {
		fmt.Println("Error running program:", err)
		os.Exit(1)
	}
}
