package tui

import (
	"fmt"

	tea "github.com/charmbracelet/bubbletea"
	"github.com/charmbracelet/lipgloss"
)

type pressAnyKeyModel struct {
	msg string
}

func (m pressAnyKeyModel) Init() tea.Cmd {
	return nil
}

func (m pressAnyKeyModel) Update(msg tea.Msg) (tea.Model, tea.Cmd) {
	switch msg.(type) {
	case tea.KeyMsg:
		return m, tea.Quit
	}
	return m, nil
}

func (m pressAnyKeyModel) View() string {
	var style = lipgloss.NewStyle().
		Bold(true).
		Blink(true)
	return style.Render(m.msg)
}

func PressAnyKey(msg ...string) error {
	message := "続行するには何かキーを押してください"
	if len(msg) != 0 {
		message = msg[0]
	}
	p := tea.NewProgram(pressAnyKeyModel{msg: message})
	if _, err := p.Run(); err != nil {
		return fmt.Errorf("cannot start PressAnyKey: %w", err)
	}
	return nil
}
