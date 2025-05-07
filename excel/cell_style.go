package excel

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

// cellStyle defines cell style flags.
type cellStyle int

const (
	// Font styles
	cellStyleNormal cellStyle = 0
	cellStyleBold             = 1 << iota

	// Fill color styles
	fillGray1     // 16 灰1 (&H808080)
	fillGray2     // 48 灰2 (&H969696)
	fillGray3     // 15 灰3 (&HC0C0C0)
	fillPink      // 26 (ピンク) CAUTION
	fillYellow    // 27 (黄色) NOTE
	fillLightBlue // 28 (薄い青) HINT

	// Thin border styles
	bT // thin border top
	bL // thin border left
	bR // thin border right
	bB // thin border bottom

	// Thick border styles
	bbT // thick border top
	bbL // thick border left
	bbR // thick border right
	bbB // thick border bottom
)

// applyCellStyle applies all cell styles to the current sheet.
func (e *Excel) applyCellStyle() error {
	for cell, style := range e.cellStyleMap {
		styleID, ok := e.cellStyleIDs[style]
		if !ok {
			return fmt.Errorf("cannot find style ID for cell '%s' in sheet '%s'",
				cell, e.sheet)
		}
		if err := e.f.SetCellStyle(e.sheet, cell, cell, styleID); err != nil {
			return fmt.Errorf(
				"failed to apply styles to cell '%s' in sheet '%s': %w",
				cell, e.sheet, err)
		}
	}
	return nil
}

// SetCellStyle applies a style to the specified cell.
func (e *Excel) SetCellStyle(cell string, style cellStyle) error {
	if style == cellStyleNormal {
		return nil
	}
	style |= e.cellStyleMap[cell]
	e.cellStyleMap[cell] = style

	if _, ok := e.cellStyleIDs[style]; ok {
		return nil
	}

	// Font
	var font excelize.Font
	if style&cellStyleBold != 0 {
		font = excelize.Font{Size: e.fontSize, Bold: true}
	}

	// Fill
	var fill excelize.Fill
	// Fill: excelize.Fill{Type: "pattern",
	// 	Color:   []string{"E0EBF5"},
	// 	Pattern: 1},
	// Ref: https://nako-itnote.com/excel-colorindex-rgb/
	if style&fillGray1 != 0 { // 灰1
		fill = excelize.Fill{
			Type: "pattern", Color: []string{"808080"}, Pattern: 1}
	}
	if style&fillGray2 != 0 { // 灰2
		fill = excelize.Fill{
			Type: "pattern", Color: []string{"969696"}, Pattern: 1}
	}
	if style&fillGray3 != 0 { // 灰3
		fill = excelize.Fill{
			Type: "pattern", Color: []string{"C0C0C0"}, Pattern: 1}
	}
	if style&fillPink != 0 { // 26 ピンク
		fill = excelize.Fill{
			Type: "pattern", Color: []string{"FF00FF"}, Pattern: 1}
	}
	if style&fillYellow != 0 { // 27 黄色
		fill = excelize.Fill{
			Type: "pattern", Color: []string{"FFFF00"}, Pattern: 1}
	}
	if style&fillLightBlue != 0 { // 28 薄い青
		fill = excelize.Fill{
			Type: "pattern", Color: []string{"#00FFFF"}, Pattern: 1}
	}

	// Border
	var border []excelize.Border

	// border thin (style=1): top, left, right, bottom
	if style&bT != 0 {
		border = append(border, excelize.Border{
			Type: "top", Style: 1, Color: "000000"})
	}
	if style&bL != 0 {
		border = append(border, excelize.Border{
			Type: "left", Style: 1, Color: "000000"})
	}
	if style&bR != 0 {
		border = append(border, excelize.Border{
			Type: "right", Style: 1, Color: "000000"})
	}
	if style&bB != 0 {
		border = append(border, excelize.Border{
			Type: "bottom", Style: 1, Color: "000000"})
	}

	// border thick (sytle=5): top, left, right, bottom
	if style&bbT != 0 {
		border = append(border, excelize.Border{
			Type: "top", Style: 5, Color: "000000"})
	}
	if style&bbL != 0 {
		border = append(border, excelize.Border{
			Type: "left", Style: 5, Color: "000000"})
	}
	if style&bbR != 0 {
		border = append(border, excelize.Border{
			Type: "right", Style: 5, Color: "000000"})
	}
	if style&bbB != 0 {
		border = append(border, excelize.Border{
			Type: "bottom", Style: 5, Color: "000000"})
	}

	id, err := e.f.NewStyle(
		&excelize.Style{
			Font:   &font,
			Fill:   fill,
			Border: border,
		})
	if err != nil {
		return fmt.Errorf("failed to initialize Bold cell style: %w", err)
	}
	e.cellStyleIDs[style] = id

	return nil
}

// SetCellStyleForCurrentCell applies the specified style to the current cell.
func (e *Excel) SetCellStyleForCurrentCell(style cellStyle) error {
	cell, err := e.Cell()
	if err != nil {
		return fmt.Errorf("failed to get cell position: %w", err)
	}
	return e.SetCellStyle(cell, style)
}

// NewStyle returns the default Normal cell style.
func NewStyle() cellStyle {
	return cellStyleNormal
}

// Bold enables the Bold flag for the cell style.
func (c cellStyle) Bold() cellStyle {
	return c | cellStyleBold
}

// Add adds cell sytles.
//
// Example:
//
//	style := NewStyle().Add(bT, bL)
func (c cellStyle) Add(styles ...cellStyle) cellStyle {
	for _, s := range styles {
		c |= s
	}
	return c
}
