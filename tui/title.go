package tui

import (
	"fmt"

	"github.com/charmbracelet/lipgloss"
)

func Title(msg string) {
	var style = lipgloss.NewStyle().
		BorderStyle(lipgloss.RoundedBorder()).
		BorderForeground(lipgloss.Color("228")).
		BorderBackground(lipgloss.Color("63")).
		BorderTop(true).
		BorderLeft(true).
		BorderRight(true).
		BorderBottom(true).
		PaddingTop(1).
		PaddingLeft(1).
		PaddingRight(1).
		PaddingBottom(1).
		Width(64).
		Foreground(lipgloss.Color("202")).
		Align(lipgloss.Center)
	fmt.Println(style.Render(msg))
}

func MsgBox(msg string) {
	var style = lipgloss.NewStyle().
		BorderStyle(lipgloss.RoundedBorder()).
		BorderForeground(lipgloss.Color("228")).
		BorderTop(true).
		BorderLeft(true).
		BorderRight(true).
		BorderBottom(true).
		Width(64).
		Foreground(lipgloss.Color("63"))
	fmt.Println(style.Render(msg))
}
