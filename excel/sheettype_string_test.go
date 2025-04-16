package excel

import "testing"

func TestSheetTypeString(t *testing.T) {
	tests := []struct {
		name     string
		typ      SheetType
		expected string
	}{
		{"Unknown", SheetTypeUnknown, "Unknown"},
		{"Normal", SheetTypeNormal, "Normal"},
		{"TOCCover", SheetTypeTOC, "TOC"},
	}

	for _, tt := range tests {
		got := tt.typ.String()
		if got != tt.expected {
			t.Errorf("%s: want %s, but got %s", tt.name, tt.expected, got)
		}
	}
}
