package tui

import (
	"errors"
	"fmt"
	"os"
	"strings"
	"time"

	"github.com/charmbracelet/bubbles/filepicker"
	tea "github.com/charmbracelet/bubbletea"
)

type filePickerModel struct {
	filepicker   filepicker.Model
	selectedFile string
	quitting     bool
	err          error
	prompt       string
}

type clearErrorMsg struct{}

func clearErrorAfter(t time.Duration) tea.Cmd {
	return tea.Tick(t, func(_ time.Time) tea.Msg {
		return clearErrorMsg{}
	})
}

func (m filePickerModel) Init() tea.Cmd {
	return m.filepicker.Init()
}

func (m filePickerModel) Update(msg tea.Msg) (tea.Model, tea.Cmd) {
	switch msg := msg.(type) {
	case tea.KeyMsg:
		switch msg.String() {
		case "ctrl+c", "q":
			m.quitting = true
			return m, tea.Quit
		}
	case clearErrorMsg:
		m.err = nil
	}

	var cmd tea.Cmd
	m.filepicker, cmd = m.filepicker.Update(msg)

	// Did the user select a file?
	if didSelect, path := m.filepicker.DidSelectFile(msg); didSelect {
		// Get the path of the selected file.
		m.selectedFile = path
		m.quitting = true
		return m, tea.Quit
	}

	// Did the user select a disabled file?
	// This is only necessary to display an error to the user.
	if didSelect, path := m.filepicker.DidSelectDisabledFile(msg); didSelect {
		// Let's clear the selectedFile and display an error.
		m.err = errors.New(path + " 拡張子が違います")
		m.selectedFile = ""
		return m, tea.Batch(cmd, clearErrorAfter(2*time.Second))
	}

	return m, cmd
}

func (m filePickerModel) View() string {
	if m.quitting {
		return ""
	}
	var s strings.Builder
	// s.WriteString("\n  ")
	s.WriteString(m.filepicker.Styles.Selected.Render(m.prompt) + "\n")
	if m.err != nil {
		s.WriteString(m.filepicker.Styles.DisabledFile.Render(m.err.Error()))
	} else if m.selectedFile == "" {
		s.WriteString("←: 親ディレクトリ・↑: 上へ・↓: 下へ・enter: 選択")
		// } else {
		// s.WriteString("選択: " + m.filepicker.Styles.Selected.Render(m.selectedFile))
	}
	s.WriteString("\n\n" + m.filepicker.View() + "\n")
	return s.String()
}

// filepicker is a file selection dialog.
func FilePicker(prompt string, extension []string) (string, error) {
	fp := filepicker.New()
	fp.AllowedTypes = extension
	// fp.CurrentDirectory, _ = os.UserHomeDir()
	fp.CurrentDirectory, _ = os.Getwd()

	m := filePickerModel{
		filepicker: fp,
		prompt:     prompt,
	}
	tm, _ := tea.NewProgram(&m).Run()
	mm := tm.(filePickerModel)
	selected := mm.selectedFile
	if len(selected) > 0 {
		fmt.Println("\n  選択: " + m.filepicker.Styles.Selected.Render(mm.selectedFile) + "\n")
		return mm.selectedFile, nil
	}
	return "", errors.New("cannot select a file")
}
