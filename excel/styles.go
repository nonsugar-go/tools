package excel

import (
	"errors"
	"fmt"

	"github.com/xuri/excelize/v2"
)

// cellStyle defines cell style flags.
type cellStyle uint64

const (
	styleNormal cellStyle = 0

	// Font sizes
	fontSize12 cellStyle = 1 << iota // 12
	fontSize20                       // 20

	// Font styles
	fontBold

	// Font color styles

	// 標準の色
	fontDeepRed    // #C00000 濃い赤
	fontRed        // #FF0000 赤
	fontOrange     // #FFC000 オレンジ
	fontYellow     // #FFFF00 黄
	fontLightGreen // #92D050 薄い緑
	fontGreen      // #00B050 緑
	fontLightBlue  // #00B0F0 薄い青
	fontBlue       // #0070C0 青
	fontDarkBlue   // #002060 濃い青
	fontPurple     // #7030A0 紫

	fontHyperLink // #0563C1 ハイパーリンク用

	// Fill color styles

	// 標準の色
	fillDeepRed    // #C00000 濃い赤
	fillRed        // #FF0000 赤
	fillOrange     // #FFC000 オレンジ
	fillYellow     // #FFFF00 黄
	fillLightGreen // #92D050 薄い緑
	fillGreen      // #00B050 緑
	fillLightBlue  // #00B0F0 薄い青
	fillBlue       // #0070C0 青
	fillDarkBlue   // #002060 濃い青
	fillPurple     // #7030A0 紫

	// グレー
	fillGray1 // #808080 グレー1
	fillGray2 // #A6A6A6 グレー2
	fillGray3 // #BFBFBF グレー3
	fillGray4 // #D9D9D9 グレー4
	fillGray5 // #F2F2F2 グレー5

	fillHeaderColor1 // #808080 <-- Excel Macro のヘッダ1
	fillHeaderColor2 // #969696 <-- Excel Macro のヘッダ2
	fillHeaderColor3 // #C0C0C0 <-- Excel Macro のヘッダ3

	fillCaution // #FF00FF ピンク <-- Excel Macro の CAUTION用
	fillNote    // #FFFF00 黄色 <-- Excel Macro のNOTE用
	fillHint    // #00FFFF 薄い青 <-- Excel Macro のHINT用

	// Alignment 配置
	alignmentHorizontalCenter // 横位置=中央揃え "center"
	alignmentShrinkToFit      // 文字の制御: 縮小して全体を表示する=true
	alignmentVerticalCenterl  // 縦位置=中央揃え "center"
	alignmentWrapText         // 文字の制御: 折り返して全体を表示する=true

	// Thin border styles
	// Index=1 Name=Continuous Weight=1
	b1L // left border
	b1T // top border
	b1R // right border
	b1B // bottom border

	// Medium border styles
	// Index=2 Name=Continuous Weight=2
	b2L // left border
	b2T // top border
	b2R // right border
	b2B // bottom border

	// Thick border styles
	// Index=5 Name=Continuous Weight=3
	b3L // left border
	b3T // top border
	b3R // right border
	b3B // bottom border

	// Double border style
	// Index=6 Name=Double Weight=3
	bdL // left border
	bdT // top border
	bdR // right border
	bdB // bottom border
)

type BorderType int

const (
	// Index | Name          | Weight | Style
	// 0     | None          | 0      |
	BorderNone BorderType = iota
	// 1     | Continuous    | 1      | -----------
	BorderContinuousWeight1
	// 2     | Continuous    | 2      | -----------
	BorderContinuousWeight2
	// 3     | Dash          | 1      | - - - - - -
	// BorderDashWeight1
	// 4     | Dot           | 1      | . . . . . .
	// BorderDotWeight1
	// 5     | Continuous    | 3      | -----------
	BorderContinuousWeight3
	// 6     | Double        | 3      | ===========
	BorderDoubleWeight3
	// 7     | Continuous    | 0      | -----------
	// BorderContinuousWeight0
	// 8     | Dash          | 2      | - - - - - -
	// BorderDashWeight2
	// 9     | Dash Dot      | 1      | - . - . - .
	// BorderDashDotWeight1
	// 10    | Dash Dot      | 2      | - . - . - .
	// BorderDashDotWeight2
	// 11    | Dash Dot Dot  | 1      | - . . - . .
	// BorderDashDotDotWeight1
	// 12    | Dash Dot Dot  | 2      | - . . - . .
	// BorderDashDotDotWeight2
	// 13    | SlantDash Dot | 2      | / - . / - .
	// BorderSlantDashDotWeight2
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

// SetStyleForCell applies a style to the specified cell.
func (e *Excel) SetStyleForCell(cell string, style cellStyle) error {
	if style == styleNormal {
		return nil
	}
	add_style := style
	style |= e.cellStyleMap[cell]

	if _, ok := e.cellStyleIDs[style]; ok {
		e.cellStyleMap[cell] = style
		return nil
	}

	// フォントのサイズについて排他処理を実施
	// 今回引数で追加したスタイルから先にチェックする

	const (
		fontSizeAll cellStyle = fontSize12 | fontSize20
	)

	style_copy := style
	if style&fontSizeAll != 0 {
		style &^= fontSizeAll
		if add_style&fontSizeAll != 0 {
			style_copy = add_style
		}
		for _, s := range []cellStyle{fontSize12, fontSize20} {
			if style_copy&s != 0 {
				style |= s
				break
			}
		}
	}

	// Font size
	fontSize := e.fontSize
	if style&fontSize12 != 0 {
		fontSize = 12
	} else if style&fontSize20 != 0 {
		fontSize = 20
	}

	// Font style
	var bold bool
	if style&fontBold != 0 {
		bold = true
	}

	// フォントの色について排他処理を実施
	// 今回引数で追加したスタイルから先にチェックする

	const (
		fontColorAll cellStyle = fontDeepRed | fontRed | fontOrange |
			fontYellow | fontLightGreen | fontGreen | fontLightBlue |
			fontBlue | fontDarkBlue | fontPurple | fontHyperLink
	)

	style_copy = style
	if style&fontColorAll != 0 {
		style &^= fontColorAll
		if add_style&fontColorAll != 0 {
			style_copy = add_style
		}
		for _, s := range []cellStyle{fontDeepRed, fontRed, fontOrange,
			fontYellow, fontLightGreen, fontGreen, fontLightBlue,
			fontBlue, fontDarkBlue, fontPurple, fontHyperLink} {
			if style_copy&s != 0 {
				style |= s
				break
			}
		}
	}

	// Font color
	var fontColor string
	if style&fontDeepRed != 0 { // 濃い赤
		fontColor = "C00000"
	} else if style&fontRed != 0 { // 赤
		fontColor = "FF0000"
	} else if style&fontOrange != 0 { // オレンジ
		fontColor = "FFC000"
	} else if style&fontYellow != 0 { // 黄
		fontColor = "FFFF00"
	} else if style&fontLightGreen != 0 { // 薄い緑
		fontColor = "92D050"
	} else if style&fontGreen != 0 { // 緑
		fontColor = "00B050"
	} else if style&fontLightBlue != 0 { // 薄い青
		fontColor = "00B0F0"
	} else if style&fontBlue != 0 { // 青
		fontColor = "0070C0"
	} else if style&fontDarkBlue != 0 { // 濃い青
		fontColor = "002060"
	} else if style&fontPurple != 0 { // 紫
		fontColor = "7030A0"
	} else if style&fontHyperLink != 0 { // ハイパーリンク用
		fontColor = "0563C1"
	}

	// 塗りつぶしについて排他処理を実施
	// 今回引数で追加したスタイルから先にチェックする

	const (
		fillAll cellStyle = fillDeepRed | fillRed | fillOrange | fillYellow |
			fillLightGreen | fillGreen | fillLightBlue | fillBlue |
			fillDarkBlue | fillPurple | fillGray1 | fillGray2 | fillGray3 |
			fillGray4 | fillGray5 | fillHeaderColor1 | fillHeaderColor2 |
			fillHeaderColor3 | fillCaution | fillNote | fillHint
	)

	style_copy = style
	if style&fillAll != 0 {
		style &^= fillAll
		if add_style&fillAll != 0 {
			style_copy = add_style
		}
		for _, s := range []cellStyle{fillDeepRed, fillRed, fillOrange,
			fillYellow, fillLightGreen, fillGreen, fillLightBlue, fillBlue,
			fillDarkBlue, fillPurple, fillGray1, fillGray2, fillGray3,
			fillGray4, fillGray5, fillHeaderColor1, fillHeaderColor2,
			fillHeaderColor3, fillCaution, fillNote, fillHint} {
			if style_copy&s != 0 {
				style |= s
				break
			}
		}
	}

	// Fill
	// Ref: https://nako-itnote.com/excel-colorindex-rgb/
	colorCode := ""

	if style&fillDeepRed != 0 { // 濃い赤
		colorCode = "C00000"
	} else if style&fillRed != 0 { // 赤
		colorCode = "FF0000"
	} else if style&fillOrange != 0 { // オレンジ
		colorCode = "FFC000"
	} else if style&fillYellow != 0 { // 黄
		colorCode = "FFFF00"
	} else if style&fillLightGreen != 0 { // 薄い緑
		colorCode = "92D050"
	} else if style&fillGreen != 0 { // 緑
		colorCode = "00B050"
	} else if style&fillLightBlue != 0 { // 薄い青
		colorCode = "00B0F0"
	} else if style&fillBlue != 0 { // 青
		colorCode = "0070C0"
	} else if style&fillDarkBlue != 0 { // 濃い青
		colorCode = "002060"
	} else if style&fillPurple != 0 { // 紫
		colorCode = "7030A0"
	} else if style&fillGray1 != 0 { // グレー1
		colorCode = "808080"
	} else if style&fillGray2 != 0 { // グレー2
		colorCode = "A6A6A6"
	} else if style&fillGray3 != 0 { // グレー3
		colorCode = "BFBFBF"
	} else if style&fillGray4 != 0 { // グレー4
		colorCode = "D9D9D9"
	} else if style&fillGray5 != 0 { // グレー5
		colorCode = "F2F2F2"
	} else if style&fillHeaderColor1 != 0 { // <-- Excel Macro のヘッダ1
		colorCode = "808080"
	} else if style&fillHeaderColor2 != 0 { // <-- Excel Macro のヘッダ2
		colorCode = "969696"
	} else if style&fillHeaderColor3 != 0 { // <-- Excel Macro のヘッダ3
		colorCode = "C0C0C0"
	} else if style&fillCaution != 0 { // ピンク <-- Excel Macro の CAUTION用
		colorCode = "FF00FF"
	} else if style&fillNote != 0 { // 黄色 <-- Excel Macro のNOTE用
		colorCode = "FFFF00"
	} else if style&fillHint != 0 { // 薄い青 <-- Excel Macro のHINT用
		colorCode = "00FFFF"
	}

	var fill excelize.Fill
	if colorCode != "" {
		fill = excelize.Fill{
			Type: "pattern", Color: []string{colorCode}, Pattern: 1}
	}

	// Alignment 配置
	var (
		varAlignmentHorizontal  = ""
		varAlignmentShrinkToFit = false
		varAlignmentVertical    = ""
		varAlignmentWrapText    = false
	)
	if style&alignmentHorizontalCenter != 0 {
		varAlignmentHorizontal = "center"
	}
	if style&alignmentShrinkToFit != 0 {
		varAlignmentShrinkToFit = true
	}
	if style&alignmentVerticalCenterl != 0 {
		varAlignmentVertical = "center"
	}
	if style&alignmentWrapText != 0 {
		varAlignmentWrapText = true
	}

	// 罫線について排他処理を実施
	// 今回引数で追加したスタイルから先にチェックする

	const (
		bLAll cellStyle = b1L | b2L | b3L | bdL
		bTAll           = b1T | b2T | b3T | bdT
		bRAll           = b1R | b2R | b3R | bdR
		bBAll           = b1B | b2B | b3B | bdB
	)

	// Border left
	style_copy = style
	if style&bLAll != 0 {
		style &^= bLAll
		if add_style&bLAll != 0 {
			style_copy = add_style
		}
		for _, s := range []cellStyle{b1L, b2L, b3L, bdL} {
			if style_copy&s != 0 {
				style |= s
				break
			}
		}
	}

	// Border top
	style_copy = style
	if style&bTAll != 0 {
		style &^= bTAll
		if add_style&bTAll != 0 {
			style_copy = add_style
		}
		for _, s := range []cellStyle{b1T, b2T, b3T, bdT} {
			if style_copy&s != 0 {
				style |= s
				break
			}
		}
	}

	// Border right
	style_copy = style
	if style&bRAll != 0 {
		style &^= bRAll
		if add_style&bRAll != 0 {
			style_copy = add_style
		}
		for _, s := range []cellStyle{b1R, b2R, b3R, bdR} {
			if style_copy&s != 0 {
				style |= s
				break
			}
		}
	}

	// Border bottom
	style_copy = style
	if style&bBAll != 0 {
		style &^= bBAll
		if add_style&bBAll != 0 {
			style_copy = add_style
		}
		for _, s := range []cellStyle{b1B, b2B, b3B, bdB} {
			if style_copy&s != 0 {
				style |= s
				break
			}
		}
	}

	// Border
	var border []excelize.Border

	for _, bs := range []struct {
		cellStyle cellStyle
		typ       string
		style     int
	}{
		{b1L, "left", 1},
		{b1T, "top", 1},
		{b1R, "right", 1},
		{b1B, "bottom", 1},

		{b2L, "left", 2},
		{b2T, "top", 2},
		{b2R, "right", 2},
		{b2B, "bottom", 2},

		{b3L, "left", 5},
		{b3T, "top", 5},
		{b3R, "right", 5},
		{b3B, "bottom", 5},

		{bdL, "left", 6},
		{bdT, "top", 6},
		{bdR, "right", 6},
		{bdB, "bottom", 6},
	} {
		if style&bs.cellStyle != 0 {
			border = append(border, excelize.Border{
				Type: bs.typ, Color: "000000", Style: bs.style})
		}
	}

	id, err := e.f.NewStyle(
		&excelize.Style{
			Font: &excelize.Font{
				Size: fontSize, Bold: bold, Color: fontColor,
			},
			Fill:   fill,
			Border: border,
			Alignment: &excelize.Alignment{
				Horizontal:      varAlignmentHorizontal,
				Indent:          0,
				JustifyLastLine: false,
				ReadingOrder:    0,
				RelativeIndent:  0,
				ShrinkToFit:     varAlignmentShrinkToFit,
				TextRotation:    0,
				Vertical:        varAlignmentVertical,
				WrapText:        varAlignmentWrapText,
			},
		})
	if err != nil {
		return fmt.Errorf("failed to initialize cell style: %w", err)
	}
	e.cellStyleIDs[style] = id
	e.cellStyleMap[cell] = style

	return nil
}

// SetStyleForCellRange applies a style to the specified cell range
func (e *Excel) SetStyleForCellRange(
	topLeftCell, bottomRightCell string, style cellStyle) error {
	hCol, hRow, err := excelize.CellNameToCoordinates(topLeftCell)
	if err != nil {
		return err
	}

	vCol, vRow, err := excelize.CellNameToCoordinates(bottomRightCell)
	if err != nil {
		return err
	}

	// Normalize the range, such correct C1:B3 to B1:C3.
	if vCol < hCol {
		vCol, hCol = hCol, vCol
	}

	if vRow < hRow {
		vRow, hRow = hRow, vRow
	}

	for r := hRow; r <= vRow; r++ {
		for c := hCol; c <= vCol; c++ {
			cell, err := excelize.CoordinatesToCellName(c, r)
			if err != nil {
				return err
			}
			if err := e.SetStyleForCell(cell, style); err != nil {
				return err
			}
		}
	}
	return nil
}

// SetStyle applies the specified style to the current cell.
func (e *Excel) SetStyle(style cellStyle) error {
	cell, err := e.Cell()
	if err != nil {
		return fmt.Errorf("failed to get cell position: %w", err)
	}
	return e.SetStyleForCell(cell, style)
}

// NewStyle combines the default Normal cell style with additional styles.
func NewStyle(styles ...cellStyle) cellStyle {
	c := styleNormal
	for _, s := range styles {
		c |= s
	}
	return c
}

// add adds cell sytles.
//
// Example:
//
//	style := NewStyle().add(bT, bL)
func (c cellStyle) add(styles ...cellStyle) cellStyle {
	for _, s := range styles {
		c |= s
	}
	return c
}

// Bold enables the Bold flag for the cell style.
//
// Example:
//
//	style := NewStyle().Bold()
func (c cellStyle) Bold() cellStyle {
	return c | fontBold
}

// DrawBorders applies borders to a specified range of cells.
//
// BoderType:
func (e *Excel) DrawBorders(topLeftCell, bottomRightCell string,
	borderType BorderType) error {
	switch borderType {
	case BorderContinuousWeight1,
		BorderContinuousWeight2,
		BorderContinuousWeight3,
		BorderDoubleWeight3:
		// 何もしない
	default:
		return fmt.Errorf("unsupported border type: %d", borderType)
	}

	hCol, hRow, err := excelize.CellNameToCoordinates(topLeftCell)
	if err != nil {
		return err
	}

	vCol, vRow, err := excelize.CellNameToCoordinates(bottomRightCell)
	if err != nil {
		return err
	}

	// Normalize the range, such correct C1:B3 to B1:C3.
	if vCol < hCol {
		vCol, hCol = hCol, vCol
	}

	if vRow < hRow {
		vRow, hRow = hRow, vRow
	}

	switch {
	case hRow == vRow && hCol == vCol:
		return errors.New("cell range must consist of multiple cells")
	case hRow == vRow:
		// Single row: draw top border only
		var style cellStyle
		switch borderType {
		case BorderContinuousWeight1:
			style = style.add(b1T)
		case BorderContinuousWeight2:
			style = style.add(b2T)
		case BorderContinuousWeight3:
			style = style.add(b3T)
		case BorderDoubleWeight3:
			style = style.add(bdT)
		}
		if err := e.SetStyleForCellRange(
			topLeftCell, bottomRightCell, style); err != nil {
			return fmt.Errorf("failed to draw borders: %w", err)
		}
	case hCol == vCol:
		// Single column: draw left border only
		var style cellStyle
		switch borderType {
		case BorderContinuousWeight1:
			style = style.add(b1L)
		case BorderContinuousWeight2:
			style = style.add(b2L)
		case BorderContinuousWeight3:
			style = style.add(b3L)
		case BorderDoubleWeight3:
			style = style.add(bdL)
		}
		if err := e.SetStyleForCellRange(
			topLeftCell, bottomRightCell, style); err != nil {
			return fmt.Errorf("failed to draw borders: %w", err)
		}
	default:
		// Multiple rows and columns: draw all borders
		for r := hRow; r <= vRow; r++ {
			for c := hCol; c <= vCol; c++ {
				cell, err := excelize.CoordinatesToCellName(c, r)
				if err != nil {
					return err
				}
				var style cellStyle
				switch r {
				case hRow:
					switch borderType {
					case BorderContinuousWeight1:
						style = style.add(b1T)
					case BorderContinuousWeight2:
						style = style.add(b2T)
					case BorderContinuousWeight3:
						style = style.add(b3T)
					case BorderDoubleWeight3:
						style = style.add(bdT)
					}
				case vRow:
					switch borderType {
					case BorderContinuousWeight1:
						style = style.add(b1B)
					case BorderContinuousWeight2:
						style = style.add(b2B)
					case BorderContinuousWeight3:
						style = style.add(b3B)
					case BorderDoubleWeight3:
						style = style.add(bdB)
					}
				}
				if err := e.SetStyleForCell(cell, style); err != nil {
					return fmt.Errorf("failed to draw borders: %w", err)
				}
				style = NewStyle()
				switch c {
				case hCol:
					switch borderType {
					case BorderContinuousWeight1:
						style = style.add(b1L)
					case BorderContinuousWeight2:
						style = style.add(b2L)
					case BorderContinuousWeight3:
						style = style.add(b3L)
					case BorderDoubleWeight3:
						style = style.add(bdL)
					}
				case vCol:
					switch borderType {
					case BorderContinuousWeight1:
						style = style.add(b1R)
					case BorderContinuousWeight2:
						style = style.add(b2R)
					case BorderContinuousWeight3:
						style = style.add(b3R)
					case BorderDoubleWeight3:
						style = style.add(bdR)
					}
				}
				if err := e.SetStyleForCell(cell, style); err != nil {
					return fmt.Errorf("failed to draw borders: %w", err)
				}
			}
		}
	}

	return nil
}
