package excel

import (
	"errors"
	"fmt"
	"strconv"
	"strings"

	"github.com/nonsugar-go/tools/excel/dataframe"
	"github.com/xuri/excelize/v2"
)

// TODO: HasVal, GetVal に置き換えらるところは、置き換える

// pageSetting configures page-specific properties.
func (e *Excel) pageSetting(sheetType SheetType, title string) error {
	if err := e.f.SetSheetProps(
		e.sheet,
		&excelize.SheetPropsOptions{
			AutoPageBreaks:                    &boolTrue,
			BaseColWidth:                      &uint80x0,
			CodeName:                          (*string)(nil),
			CustomHeight:                      &boolTrue,
			DefaultColWidth:                   &defaultColWidth,
			DefaultRowHeight:                  &defaultRowHeight,
			EnableFormatConditionsCalculation: &boolTrue,
			FitToPage:                         (*bool)(nil),
			OutlineSummaryBelow:               &boolTrue,
			OutlineSummaryRight:               (*bool)(nil),
			Published:                         &boolTrue,
			TabColorIndexed:                   (*int)(nil),
			TabColorRGB:                       (*string)(nil),
			TabColorTheme:                     (*int)(nil),
			TabColorTint:                      (*float64)(nil),
			ThickBottom:                       &boolFalse,
			ThickTop:                          &boolFalse,
			ZeroHeight:                        &boolFalse,
		},
	); err != nil {
		return fmt.Errorf("cannot set sheet props: %s: %w", e.sheet, err)
	}

	// タイトルの設定
	switch sheetType {
	case SheetTypeNormal:
		e.Col, e.Row = 1, 1
		if err := e.f.SetCellStr(e.sheet, "A1", title); err != nil {
			return fmt.Errorf(
				"failed to set value for cell 'A1' on sheet '%s': %s: %w",
				e.sheet, title, err)
		}
		if err := e.MarkHeader(1); err != nil {
			return err
		}
	case SheetTypeTOC:
		e.Col, e.Row = 1, 1
		title = "目次"
		if err := e.f.SetCellStr(e.sheet, "A1", title); err != nil {
			return fmt.Errorf(
				"failed to set value for cell 'A1' on sheet '%s': %s: %w",
				e.sheet, title, err)
		}
	case SheetTypeCover:
		e.Col, e.Row = 1, 6
		if err := e.f.SetCellStr(e.sheet, "A6", title); err != nil {
			return fmt.Errorf(
				"failed to set value for cell 'A6' on sheet '%s': %s: %w",
				e.sheet, title, err)
		}
		if err := e.f.MergeCell(e.sheet, "A6", maxRightCell+"11"); err != nil {
			return fmt.Errorf(
				"failed to merge cells for 'A6:%s11' on sheet '%s': %w",
				maxRightCell, e.sheet, err)
		}
		// TODO:
		if err := e.SetStyle(NewStyle(fontSize20, fontBold,
			// 配置
			alignmentHorizontalCenter, // 横位置=中央揃え
			alignmentVerticalCenterl,  // 縦位置=中央揃え
			alignmentWrapText,         // 文字の制御: 折り返して全体を表示する
		)); err != nil {
			return err
		}
		for _, v := range []struct{ cell, value string }{
			{"A20", "更新日付"},
			{"E20", "内容"},
			{"T20", "更新箇所"},
			{"AD20", "更新者"},
		} {
			if err := e.f.SetCellStr(e.sheet, v.cell, v.value); err != nil {
				return err
			}
		}
		if err := e.DrawBorders2("A20", maxRightCell+"40",
			TBorderHHeader); err != nil {
			return err
		}
	}

	switch sheetType {
	case SheetTypeNormal, SheetTypeTOC:
		e.Col, e.Row = 1, 1
		if err := e.SetStyle(NewStyle(fontSize12, fontBold)); err != nil {
			return err
		}
		for r, h := range map[int]float64{1: 15.75, 2: 3, 3: 15.75} {
			if err := e.f.SetRowHeight(e.sheet, r, h); err != nil {
				return err
			}
		}

		// 3行目の上にラインを引く
		// TOMATO Macro: 2行目の上にラインを引く
		e.DrawBorders("A3", maxRightCell+"3", BorderContinuousWeight3)

		// 印刷タイトル - タイトル行: $1:$3
		if err := e.f.SetDefinedName(&excelize.DefinedName{
			Name:     "_xlnm.Print_Titles",
			RefersTo: fmt.Sprintf("'%s'!$1:$3", e.sheet),
			Scope:    e.sheet,
		}); err != nil {
			return fmt.Errorf("failed to set print titles on sheet '%s': %w",
				e.sheet, err)
		}

	}

	// ヘッダーとフッターの設定
	if err := e.f.SetHeaderFooter(e.sheet, &excelize.HeaderFooterOptions{
		AlignWithMargins: (*bool)(nil),
		DifferentFirst:   false,
		DifferentOddEven: false,
		EvenFooter:       "",
		EvenHeader:       "",
		FirstFooter:      "",
		FirstHeader:      "",
		OddFooter:        "&C&P / &N",
		OddHeader:        "",
		ScaleWithDoc:     (*bool)(nil),
	}); err != nil {
		return fmt.Errorf("failed to set header and footer on sheet '%s': %w",
			e.sheet, err)
	}

	switch sheetType {
	case SheetTypeGridA3Landscape, SheetTypeGridA4Landscape:
		// 印刷向きが横の場合
		// TODO
	default:
		// 印刷向きが縦の場合

		// ページレイアウトの設定
		var (
			adjustTo        uint = 100          // 拡大率=100%
			blackAndWhite        = false        // 白黒印刷しない
			firstPageNumber      = (*uint)(nil) // 先頭ページ番号=自動設定
			fitToHeight          = (*int)(nil)
			fitToWidth           = (*int)(nil)
			orientation          = "portrait" // 印刷の向き=縦
			size                 = 9          // 用紙サイズ=A4 (210 mm × 297 mm)
		)
		if err := e.f.SetPageLayout(e.sheet, &excelize.PageLayoutOptions{
			AdjustTo:        &adjustTo,
			BlackAndWhite:   &blackAndWhite,
			FirstPageNumber: firstPageNumber,
			FitToHeight:     fitToHeight,
			FitToWidth:      fitToWidth,
			Orientation:     &orientation,
			Size:            &size,
		}); err != nil {
			return fmt.Errorf("failed to page layout on sheet '%s': %w",
				e.sheet, err)
		}

		// 印刷マージンの設定
		var (
			bottom       = 0.629921269229078
			footer       = 0.2362204818275031
			header       = 0.2362204818275031
			horizontally = (*bool)(nil)
			left         = 0.629921269229078
			right        = 0.2362204818275031
			top          = 0.629921269229078
			vertically   = (*bool)(nil)
		)
		if err := e.f.SetPageMargins(e.sheet,
			&excelize.PageLayoutMarginsOptions{
				Bottom:       &bottom,
				Footer:       &footer,
				Header:       &header,
				Horizontally: horizontally,
				Left:         &left,
				Right:        &right,
				Top:          &top,
				Vertically:   vertically,
			}); err != nil {
			return fmt.Errorf("failed to page layout margins on sheet '%s': %w",
				e.sheet, err)
		}
	}

	switch sheetType {
	case SheetTypeNormal:
		e.Col, e.Row = 1, 3
	default:
		e.Col, e.Row = 1, 1
	}

	return nil
}

type TomatoBorderType int

const (
	TBorderUnknown TomatoBorderType = iota

	TBorderNested    // "0-9" 入れ子構造の表
	TBorderHHeader   // "H" 水平ヘッダのある表
	TBorderHHeaderG  // "G" 水平ヘッダのある表 (グループ対応)
	TBorderVHeader   // "V" 垂直ヘッダのある表
	TBorderCaution   // "C" 警告用罫線
	TBorderNote      // "N" 注意用罫線
	TBorderInfo      // "I" ヒント用罫線
	TBorderCode      // "T" TeleType文字 (コード部分に使用)
	TBorderBullet    // "B" bullet 「■□●○」だけのセルに次のセルの内容を結合する
	TBorderSchedule  // "s" 手順書用の枠 (項番の例: 1.1)
	TBorderScheduleS // "S (Shift + S) ... 手順書用の枠 (項番の例: 1)
)

// DrawBorders2 is a convenient method for easily creating tables.
func (e *Excel) DrawBorders2(topLeftCell, bottomRightCell string,
	borderType TomatoBorderType) error {
	c1, r1, err := excelize.CellNameToCoordinates(topLeftCell)
	if err != nil {
		return err
	}
	c2, r2, err := excelize.CellNameToCoordinates(bottomRightCell)
	if err != nil {
		return err
	}
	if c1 > c2 || r1 > r2 {
		return fmt.Errorf(
			"invalid range: top-left (%s) must be less than or equal to bottom-right (%s)",
			topLeftCell, bottomRightCell)
	}
	if (c2-c1+1)*(r2-r1+1) < 2 {
		return errors.New(
			"invalid range: the range must contain at least 2 cells")
	}

	switch borderType {
	case TBorderNested: // 入れ子構造の表
		return e.paramBorders(c1, r1, c2, r2)
	case TBorderHHeader:
		return e.headerBorders(borderType, c1, r1, c2, r2)
	case TBorderHHeaderG:
		return e.headerBorders(borderType, c1, r1, c2, r2)
	case TBorderVHeader:
		return e.headerBorders(borderType, c1, r1, c2, r2)
	case TBorderCaution:
		return e.admonitionBorders(borderType, c1, r1, c2, r2)
	case TBorderNote:
		return e.admonitionBorders(borderType, c1, r1, c2, r2)
	case TBorderInfo:
		return e.admonitionBorders(borderType, c1, r1, c2, r2)
	case TBorderCode:
		// TODO:
		// ' TeleType文字(等幅フォント)にし、罫線で囲う
		// If Range(cell1, cell2).Address <> originalSelection.Address Then
		//     ' 予め選択範囲を決めていない場合は、上下1行を空行と想定し範囲を広げて加工
		//     Set cell1 = cell1.Offset(-1, 0)
		//     Set cell2 = cell2.Offset(1, 0)
		//     Range(cell1, cell2).Select
		// End If
		// teleTypeBorders
	case TBorderBullet:
		// TODO:
		// ElseIf InStr(1, "BbＢｂ", Left$(msgResult, 1), vbTextCompare) <> 0 Then
		//     '「■□●○」だけのセルに次のセルの内容を結合する
		//     mergeCheckBox
	case TBorderSchedule:
		// TODO:
		// ElseIf InStr(1, "sｓ", Left$(msgResult, 1), vbBinaryCompare) <> 0 Then
		//     drawTejunsyoBorder 2
	case TBorderScheduleS:
		// TODO:
		// ElseIf InStr(1, "SＳ", Left$(msgResult, 1), vbBinaryCompare) <> 0 Then
		//     drawTejunsyoBorder 1
	default:
		return fmt.Errorf("invalid border type: %v", borderType)
	}

	return nil
}

// setStyleBorders applies a common style to all cells
// within the specified range.
func (e *Excel) setStyleBorders(col1, row1, col2, row2 int) error {
	// 全てのセルに適用するスタイルを設定する
	cell1, err := excelize.CoordinatesToCellName(col1, row1)
	if err != nil {
		return err
	}
	cell2, err := excelize.CoordinatesToCellName(col2, row2)
	if err != nil {
		return err
	}
	if err := e.SetStyleForCellRange(cell1, cell2, NewStyle(
		// 文字の配置 > 縦位置: 中央揃え
		alignmentVerticalCenterl,
		// 文字の制御 > 縮小して全体を表示する: true
		alignmentShrinkToFit,
	)); err != nil {
		return err
	}
	return nil
}

func (e *Excel) drawOuterBorders(col1, row1, col2, row2 int) error {
	// 外枠を太線で引く
	cell1, err := excelize.CoordinatesToCellName(col1, row1)
	if err != nil {
		return err
	}
	cell2, err := excelize.CoordinatesToCellName(col2, row2)
	if err != nil {
		return err
	}
	if err := e.DrawBorders(
		cell1, cell2, BorderContinuousWeight2); err != nil {
		return err
	}
	return nil
}

// headerBorders applies borders around the specified area.
// The border style is determined by the given border type.
//
// Parameters:
//
//	borderType: Specifies the style of the border.
//	col1, row1: Coordinates of the top-left corner.
//	col2, row2: Coordinates of the bottom-right corner.
func (e *Excel) headerBorders(borderType TomatoBorderType,
	col1, row1, col2, row2 int) error {
	//     0        1           2
	//     23456 78901 23456 789012 3
	//     *     *     *     *      *<-- sepCol = []int{2, 7, 12, 17, 23}
	//    ##########################
	// 07*#ID   |SRC  |DST  |ACT   #
	//    #========================#
	// 08*#A0001|host1|ALL  |PERMIT#
	//    #-----+-----+-----+------#
	// 09 #     |host2|     |      #
	//    #-----+-----+-----+------#
	// 10*#A0002|     |web1 |DENY  #
	//    #-----+-----+-----+------#
	// 11 #     |     |web2 |      #
	//    ##########################
	// 12*
	//   ^
	//   |
	//   +-- sepRow = []int{07, 08, 10, 12}
	sepCol := make([]int, 0, 2) // 列区切りとなるセルの位置
	sepRow := make([]int, 0, 2) // 行区切りとなるセルの位置

	sepCol = append(sepCol, col1)
	for c := col1 + 1; c <= col2; c++ {
		cell, err := excelize.CoordinatesToCellName(c, row1)
		if err != nil {
			return err
		}
		// 1行目かつ何か値がある場合
		// 列区切りとなるセルの位置を覚える
		value, err := e.f.GetCellValue(e.sheet, cell)
		if err != nil {
			return err
		}
		if value != "" {
			sepCol = append(sepCol, c)
		}
	} // for c
	sepCol = append(sepCol, col2+1)

	sepRow = append(sepRow, row1)
	for r := row1 + 1; r <= row2; r++ {
		if borderType == TBorderHHeaderG {
			cell, err := excelize.CoordinatesToCellName(col1, r)
			if err != nil {
				return err
			}
			// 1列目かつ何か値がある場合
			// 行区切りとなるセルの位置を覚える
			value, err := e.f.GetCellValue(e.sheet, cell)
			if err != nil {
				return err
			}
			if value != "" {
				sepRow = append(sepRow, r)
			}
		} else { // if borderType == TBorderHHeaderG
			sepRow = append(sepRow, r)
		} // if borderType == TBorderHHeaderG
	} // for r
	sepRow = append(sepRow, row2+1)

	// 全てのセルに適用するスタイルを設定する
	if err := e.setStyleBorders(col1, row1, col2, row2); err != nil {
		return err
	}
	/*
		cell1, err := excelize.CoordinatesToCellName(col1, row1)
		if err != nil {
			return err
		}
		cell2, err := excelize.CoordinatesToCellName(col2, row2)
		if err != nil {
			return err
		}
		if err := e.SetStyleForCellRange(cell1, cell2, NewStyle(
			// 文字の配置 > 縦位置: 中央揃え
			alignmentVerticalCenterl,
			// 文字の制御 > 縮小して全体を表示する: true
			alignmentShrinkToFit,
		)); err != nil {
			return err
		}
	*/

	// 列区切りとなるセルの位置でマージする
	for r := row1; r <= row2; r++ {
		prevC := sepCol[0]
		for _, c := range sepCol[1:] {
			// セルの結合
			cell1, err := excelize.CoordinatesToCellName(prevC, r)
			if err != nil {
				return err
			}
			cell2, err := excelize.CoordinatesToCellName(c-1, r)
			if err != nil {
				return err
			}
			if err := e.f.MergeCell(
				e.sheet, cell1, cell2); err != nil {
				return err
			}
			prevC = c
		} // for _, c := range sepCol
	} // for r

	// 列区切りとなるセルの位置で内側 (左) の線を引く
	for _, c := range sepCol[1 : len(sepCol)-1] {
		// 内側 (左) の線を引く
		cell1, err := excelize.CoordinatesToCellName(c, row1)
		if err != nil {
			return err
		}
		cell2, err := excelize.CoordinatesToCellName(c, row2)
		if err != nil {
			return err
		}
		if err := e.DrawBorders(
			cell1, cell2, BorderContinuousWeight1); err != nil {
			return err
		}
	} // for _, c := range sepCol[1:]

	// 内側 (上) の線を引く
	skipRows := 2
	if borderType == TBorderVHeader { // 垂直の場合
		skipRows = 1
	}
	for _, r := range sepRow[skipRows:] {
		cell1, err := excelize.CoordinatesToCellName(col1, r)
		if err != nil {
			return err
		}
		cell2, err := excelize.CoordinatesToCellName(col2, r)
		if err != nil {
			return err
		}
		if err := e.DrawBorders(
			cell1, cell2, BorderContinuousWeight1); err != nil {
			return err
		}
	} // for _, r := range sepRow[skipRows:]

	// 「水平ヘッダのある表 (グループ対応)」 の場合
	// 複数行のマージと区切りの点線を引く
	//
	//     0        1           2
	//     23456 78901 23456 789012 3
	//     *     *     *     *      *
	//    ##########################
	// 07*# ID  | SRC | DST | ACT  #
	//    #========================#
	// 08*#     |host1|     |      #
	//    #A0001+.....+ALL  +PERMIT#
	// 09 #     |host2|     |      #
	//    #-----+-----+-----+------#
	// 10*#     |     |web1 |      #
	//    #A0002+ALL  +.....+DENY  #
	// 11 #     |     |web2 |      #
	//    ##########################
	// 12*
	if borderType == TBorderHHeaderG {
		for i := 1; i < len(sepRow)-1; i++ { // ヘッダと表の次の行を除く
			// グループごとに処理 (e.g., ID=A0001, ID=A0002)
			if sepRow[i+1]-sepRow[i] < 2 { // 2行以上なければスキップ
				continue
			}
			for j := range len(sepCol) - 1 { // 表の次の列を除く
				// グループごと、列ごとに処理
				c := sepCol[j]

				// 対象の列が配列形式か？
				isArray := false
				for r := sepRow[i] + 1; r < sepRow[i+1]; r++ {
					cell, err := excelize.CoordinatesToCellName(c, r)
					if err != nil {
						return err
					}
					value, err := e.f.GetCellValue(e.sheet, cell)
					if err != nil {
						return err
					}
					if value != "" {
						isArray = true
						break
					}
				} // for r
				if !isArray {
					// 配列形式でなければ、その列をマージ
					cell1, err := excelize.CoordinatesToCellName(c, sepRow[i])
					if err != nil {
						return err
					}
					cell2, err := excelize.CoordinatesToCellName(
						sepCol[j+1]-1, sepRow[i+1]-1)
					if err != nil {
						return err
					}
					if err := e.f.MergeCell(
						e.sheet, cell1, cell2); err != nil {
						return err
					}
				} else { // if !isArray
					// 配列形式なら、内側 (上) の破線を引く
					for r := sepRow[i] + 1; r < sepRow[i+1]; r++ {
						cell1, err := excelize.CoordinatesToCellName(
							c, r)
						if err != nil {
							return err
						}
						cell2, err := excelize.CoordinatesToCellName(
							sepCol[j+1]-1, r)
						if err != nil {
							return err
						}
						if err := e.DrawBorders(
							cell1, cell2, BorderDashWeight1); err != nil {
							return err
						}
					} // for r
				} // if !isArray
			} // for j
		} // for i
	} // if borderType == TBorderHHeaderG

	// ヘッダのスタイル
	switch borderType {
	case TBorderHHeader, TBorderHHeaderG: // 水平の場合
		cell1, err := excelize.CoordinatesToCellName(col1, row1)
		if err != nil {
			return err
		}
		cell2, err := excelize.CoordinatesToCellName(col2, row1)
		if err != nil {
			return err
		}

		if err := e.SetStyleForCellRange(cell1, cell2, NewStyle(
			fillHeaderColor3,          // ヘッダに色を塗る
			alignmentHorizontalCenter, // 文字の配置 > 横位置: 中央揃え
		)); err != nil {
			return err
		}

		cell1, err = excelize.CoordinatesToCellName(col1, row1+1)
		if err != nil {
			return err
		}
		cell2, err = excelize.CoordinatesToCellName(col2, row1+1)
		if err != nil {
			return err
		}

		// ヘッダの行の上に二重線を引く
		if err := e.DrawBorders(
			cell1, cell2, BorderDoubleWeight3); err != nil {
			return err
		}
	case TBorderVHeader: // 垂直の場合
		col3 := sepCol[1]
		cell1, err := excelize.CoordinatesToCellName(col1, row1)
		if err != nil {
			return err
		}
		cell2, err := excelize.CoordinatesToCellName(col3-1, row2)
		if err != nil {
			return err
		}

		if err := e.SetStyleForCellRange(cell1, cell2, NewStyle(
			fillHeaderColor3,          // ヘッダに色を塗る
			alignmentHorizontalCenter, // 文字の配置 > 横位置: 中央揃え
		)); err != nil {
			return err
		}

		cell1, err = excelize.CoordinatesToCellName(col3, row1)
		if err != nil {
			return err
		}
		cell2, err = excelize.CoordinatesToCellName(col3, row2)
		if err != nil {
			return err
		}

		// ヘッダの行の右に二重線を引く
		if err := e.DrawBorders(
			cell1, cell2, BorderDoubleWeight3); err != nil {
			return err
		}
	default:
		panic("invalid TomatoBorderType")
	}

	// 外枠を太線で引く
	if err := e.drawOuterBorders(col1, row1, col2, row2); err != nil {
		return nil
	}
	/*
		cell1, err := excelize.CoordinatesToCellName(col1, row1)
		if err != nil {
			return err
		}
		cell2, err := excelize.CoordinatesToCellName(col2, row2)
		if err != nil {
			return err
		}
		if err := e.DrawBorders(
			cell1, cell2, BorderContinuousWeight2); err != nil {
			return err
		}
	*/

	return nil
}

// paramBorders draws nested borders for a specified cell range.
func (e *Excel) paramBorders(col1, row1, col2, row2 int) error {
	// TODO: 見直しが必要

	// 全てのセルに適用するスタイルを設定する
	if err := e.setStyleBorders(col1, row1, col2, row2); err != nil {
		return err
	}

	// 各レベルの塗りつぶしの色
	levelColor := [...]cellStyle{styleNormal,
		fillGray1, fillGray2, fillGray3, // level=1-3
		fillGray1, fillGray2, fillGray3, // level=4-6
		fillGray1, fillGray2, fillGray3} // level=7-9

	// Excel Macro: calcLevels Selection, levels
	// レベルを計算する
	// レベルとは、各行の最初に値が存在する列の位置
	// ただし、直前のレベルからひとつずつしか増えない
	// level=1 かつ 下の行が level=1 の場合は、level=0 と定義する
	// level=1 かつ 最終行の場合は、level=0 と定義
	//
	// 1-2-3-4-5
	// A         level=0 下の行が level=1 なので、level=0 と定義
	// B     b b level=1 ヘッダ行
	//   C   c c level=2
	//   D     d level=2
	//     E e   level=3
	//   F     f level=2
	// G     g   level=0 最終行なので、level=0 と定義
	levels := make(map[int]int, row2-row1+1) // 各行のレベル
	prevLevel := 1
	for r := row1; r <= row2; r++ {
		level := 0
		levels[r] = prevLevel
		for c := col1; c <= col2; c++ {
			level++
			// 何か書き込んであるセルを見つけたら
			hasVal, err := e.HasVal(c, r)
			if err != nil {
				return err
			}
			if hasVal {
				if level <= prevLevel+1 {
					// レベルはひとつずつしか上がらない
					levels[r] = level
					prevLevel = level
					break // for c
				}
			}
		} // for c
	} // for r

	// level=1 かつ 下の行が level=1 の場合は、level=0 と定義する
	// level=1 かつ 最終行の場合は、level=0 と定義
	for r := row1; r <= row2; r++ {
		if levels[r] == 1 {
			if r == row2 {
				levels[r] = 0
			} else if levels[r+1] == 1 {
				levels[r] = 0
			}
		}
	} // for r

	// ヘッダ行かどうかを返す
	isHeaderRow := func(row int, levels map[int]int) bool {
		if row >= row2 {
			// 最終行はヘッダ行でない
			return false
		}
		if levels[row] == 0 {
			// レベル0はヘッダではない
			return false
		}
		if levels[row] < levels[row+1] {
			// レベルが、次の行のレベルより小さければヘッダ行
			return true
		}
		// それ以外はヘッダ行でない
		return false
	}

	// Excel Macro: calcHeaderLevels levels, headerLevels
	// ヘッダレベルを計算する
	// ヘッダレベルとは、直近の親になるヘッダ行のレベルのこと
	headerLevels := make(map[int]int, row2-row1+1) // 各行のヘッダレベル

	for r := row1; r <= row2; r++ {
		if isHeaderRow(r, levels) {
			// ヘッダ行なら、レベルと同じ
			headerLevels[r] = levels[r]
		} else {
			// ヘッダ行でないなら、レベル - 1
			if levels[r] >= 1 {
				headerLevels[r] = levels[r] - 1
			} else {
				// level=0 なら 0
				headerLevels[r] = 0
			}
		}
	} // for r

	for r := row1; r <= row2; r++ {
		cell1Col := col1     // 結合のために topleft の列番号を記録
		isTopBorder := false // 罫線 (上) が必要なら true
		for c := col1; c <= col2; c++ {
			nColumns := c - col1 + 1
			cell, err := excelize.CoordinatesToCellName(c, r)
			if err != nil {
				return err
			}
			hasVal, err := e.HasVal(c, r)
			if err != nil {
				return err
			}
			if nColumns <= levels[r] || hasVal {
				// 列数がレベルに達していないか、何か値があるとき
				if c > col1 {
					// TODO: 左のセルが "■□●○" でない場合、
					// 左に縦線を実線で引く
					if err := e.SetStyleForCell(
						cell, NewStyle(b1L)); err != nil {
						return err
					}
				}

				// セルをマージ
				if cell1Col+1 <= c {
					cell1, err := excelize.CoordinatesToCellName(cell1Col, r)
					if err != nil {
						return err
					}
					cell2, err := excelize.CoordinatesToCellName(c-1, r)
					if err != nil {
						return err
					}
					err = e.f.MergeCell(e.sheet, cell1, cell2)
					if err != nil {
						return err
					}
				} // if cell1Col+1 <= c
				cell1Col = c
			} // if nColumns <= levels[r] || hasVal
			if isTopBorder || hasVal {
				// セルに値が存在したら、罫線 (上) を実線で引く
				if err := e.SetStyleForCell(cell, NewStyle(b1T)); err != nil {
					return err
				}
				isTopBorder = true
			}
			// 各レベルのヘッダに色を塗る
			if isHeaderRow(r, levels) && nColumns > headerLevels[r] {
				err := e.SetStyleForCell(
					cell, NewStyle(levelColor[headerLevels[r]]))
				if err != nil {
					return err
				}
			} else if nColumns <= headerLevels[r] { // if isHeaderRow
				err := e.SetStyleForCell(
					cell, NewStyle(levelColor[nColumns]))
				if err != nil {
					return err
				}
			} // if isHeaderRow
			// 最終列ならセルをマージ
			if c == col2 && cell1Col <= c {
				cell1, err := excelize.CoordinatesToCellName(cell1Col, r)
				if err != nil {
					return err
				}
				err = e.f.MergeCell(e.sheet, cell1, cell)
				if err != nil {
					return err
				}
			}
		} // for c
	} // for r

	// 外枠を太線で引く
	if err := e.drawOuterBorders(col1, row1, col2, row2); err != nil {
		return nil
	}
	return nil
}

// admonitionBorders applies borders to a specified cell range
// for different types of admonitions.
// Supports warning, note, and info styles with customizable border types.
func (e *Excel) admonitionBorders(borderType TomatoBorderType,
	col1, row1, col2, row2 int) error {
	// TODO: 精査する

	// 全てのセルに適用するスタイルを設定する
	if err := e.setStyleBorders(col1, row1, col2, row2); err != nil {
		return err
	}

	// 各行をマージ
	for r := row1; r <= row2; r++ {
		cell1, err := excelize.CoordinatesToCellName(col1, r)
		if err != nil {
			return err
		}
		cell2, err := excelize.CoordinatesToCellName(col2, r)
		if err != nil {
			return err
		}
		err = e.f.MergeCell(e.sheet, cell1, cell2)
		if err != nil {
			return err
		}
	} // for r

	// 1行目のヘッダを作成
	cell1, err := excelize.CoordinatesToCellName(col1, row1)
	if err != nil {
		return err
	}
	var value string
	var color cellStyle
	switch borderType {
	case TBorderCaution:
		value = "警告:"
		color = fillCaution
	case TBorderNote:
		value = "注意:"
		color = fillNote
	case TBorderInfo:
		value = "ヒント:"
		color = fillHint
	default:
		return fmt.Errorf(
			"failed to draw admonition borders: unsupported border type '%v'",
			borderType)
	}
	if err := e.f.SetCellStr(e.sheet, cell1, value); err != nil {
		return err
	}
	if err := e.SetStyleForCell(cell1, NewStyle(fontBold)); err != nil {
		return err
	}

	// 1行目に背景色を付ける
	cell2, err := excelize.CoordinatesToCellName(col2, row1)
	if err != nil {
		return err
	}
	if err := e.SetStyleForCellRange(
		cell1, cell2, NewStyle(color)); err != nil {
		return err
	}

	// 外枠を細線で引く
	cell2, err = excelize.CoordinatesToCellName(col2, row2)
	if err != nil {
		return err
	}
	if err := e.DrawBorders(
		cell1, cell2, BorderContinuousWeight1); err != nil {
		return err
	}

	return nil
}

// MakeTOC generates a table of contents for a document.
func (e *Excel) MakeTOC() error {
	type headersInfo struct {
		value         string // ヘッダのセルの値
		level         int    // ヘッダのレベル
		cellOfHeaders string // ヘッダのセル座標
	}
	var (
		number  [maxHeaderLevel + 1]int // ヘッダレベルごとの番号を保持
		headers []headersInfo           // ヘッダの情報
	)
	for _, sheet := range e.f.GetSheetList() {
		comments, err := e.GetSortedComments(sheet)
		if err != nil {
			return fmt.Errorf(
				"failed to retrieve sorted comments from sheet '%s': %w",
				sheet, err)
		}
		var sb strings.Builder
	OUTER:
		for _, comment := range comments {
			cell := comment.Cell
			cellAbs, err := RelCellNameToAbsCellName(cell)
			if err != nil {
				return fmt.Errorf(
					"failed to convert absolute reference for cell: %w", err)
			}
			var text string // コメントのテキスト
			isFound := false
			for _, paragraph := range comment.Paragraph {
				s := paragraph.Text
				if strings.HasPrefix(s, headerMark) {
					text = s
					isFound = true
					break
				}
			}
			if !isFound {
				continue OUTER
			}
			headerLevel, err := strconv.Atoi(
				strings.TrimLeft(text, headerMark))
			if err != nil {
				return fmt.Errorf(
					"failed to convert trimmed header level '%s' to integer: %w",
					text, err)
			}
			if headerLevel < 1 && headerLevel > maxHeaderLevel {
				return fmt.Errorf(
					"failed to validate header level: got '%d', but expected range is 1 to %d in comment '%s': %w",
					headerLevel, maxHeaderLevel, text, err)
			}
			number[headerLevel]++
			sb.Reset() // ヘッダ番号 (例: 1.2.3.)
			for i := 1; i <= headerLevel; i++ {
				sb.WriteString(strconv.Itoa(number[i]))
				sb.WriteString(".")
			}
			for i := headerLevel + 1; i <= maxHeaderLevel; i++ {
				number[i] = 0
			}
			cellValue, err := e.f.GetCellValue(sheet, cell)

			// セルの値の先頭に既に番号 (例: 1.2.3.) が含まれる場合は削除する。
			// その後、求めた番号を先頭に付与する。
			const removeChs = "0123456789.０１２３４５６７８９．"
			runeVal := []rune(cellValue)
			start := 0
			for start < len(runeVal) &&
				strings.ContainsRune(removeChs, runeVal[start]) {
				start++
			}
			sb.WriteString(strings.TrimSpace(string(runeVal[start:])))
			headerCellValue := sb.String()
			e.f.SetCellStr(sheet, cell, headerCellValue)
			headers = append(headers, headersInfo{
				value: headerCellValue,
				level: headerLevel,
				// 'sheet'!$A$1
				cellOfHeaders: fmt.Sprintf(`'%s'!%s`, sheet, cellAbs),
			})
		}
	}

	// 目次を作成する
	var cell string
	for i, header := range headers {
		// レベル1 のヘッダの前で空行を挿入
		if i != 0 {
			if header.level == 1 {
				e.LF()
			}
			e.LF()
		}
		if err := e.CR(header.level + 1).
			SetVal(header.value); err != nil {
			return fmt.Errorf("failed to set cell value in cell '%s': %w",
				cell, err)
		}
		cell, err := e.Cell()
		if err != nil {
			return fmt.Errorf(
				"failed to get current cell position: %w", err)
		}
		if err := e.f.SetCellHyperLink(
			e.sheet, cell, header.cellOfHeaders, "Location"); err != nil {
			return fmt.Errorf(
				"failed to set hyper link in cell '%s' to target '%s': %w",
				cell, header.cellOfHeaders, err)
		}
		if err := e.SetStyle(
			NewStyle(fontBold, fontHyperLink)); err != nil {
			return fmt.Errorf("failed to set cell Style for cell '%s: %w",
				cell, err)
		}
		if i == 0 {
			// 最初のヘッダのみ､開始コメントを付ける｡
			if err := e.AddComment(beginTableOfContents); err != nil {
				return fmt.Errorf("failed to add comment in cell '%s': %w",
					cell, err)
			}
		}
	}
	// 最終行の場合のみ、終了コメントを付ける。
	if len(headers) > 1 {
		if err := e.CR(2).AddComment(endTableOfContents); err != nil {
			return fmt.Errorf("failed to add comment in cell '%s': %w",
				cell, err)
		}
	}
	return nil
}

/*
       For i = 0 To NumOfHeaders - 1
           ' ヘッダをセルに入れる。
           ActiveCell.EntireRow.Insert
           ActiveSheet.Hyperlinks.Add _
               Anchor:=ActiveCell.Offset(0, headerLevels(i) - 1), _
               Address:="", SubAddress:=cellOfHeaders(i), _
               TextToDisplay:=headers(i)
           ActiveCell.Offset(0, headerLevels(i) - 1).Font.Size = 10
           ActiveCell.Offset(0, headerLevels(i) - 1).Font.Bold = True
           ActiveCell.Offset(0, headerLevels(i) - 1).Font.Underline = False

           If i = 0 Then
               ' 最初のヘッダのみ､開始コメントを付ける｡
               With ActiveCell
                   .ClearComments
                   .AddComment
                   .comment.Visible = False
                   .comment.Text Text:=BEGIN_TABLE_OF_CONTENTS
               End With
               ' 名前を定義する。
               ' ActiveWorkbook.Names.Add Name:=TOC_NAME, RefersTo:=ActiveCell
           End If

           ActiveCell.Offset(1, 0).Activate

           If i = NumOfHeaders - 1 Then
               ' 最終行の場合のみ、終了コメントを付ける。
               ActiveCell.Offset(-1, 0).Activate
               With ActiveCell
                   .ClearComments
                   .AddComment
                   .comment.Visible = False
                   .comment.Text Text:=END_TABLE_OF_CONTENTS
               End With
           End If
       Next i
   End If

*/

// MarkHeader writes header markings as comments.
func (e *Excel) MarkHeader(headerLevel int) error {
	if headerLevel < 0 || headerLevel > maxHeaderLevel {
		return fmt.Errorf("header level must be between 0 and %d, but got %d",
			maxHeaderLevel, headerLevel)
	}
	cell, err := e.Cell()
	if err != nil {
		return err
	}
	if headerLevel == 0 {
		// コメントを削除する
		err := e.f.DeleteComment(e.sheet, cell)
		if err != nil {
			return fmt.Errorf("failed to delete header mark: %w", err)
		}
	} else {
		// ヘッダの印を付ける
		if err := e.SetStyle(NewStyle(fontBold)); err != nil {
			return fmt.Errorf("failed to set header mark: %w", err)
		}
		if err := e.AddComment(
			fmt.Sprintf("%s%d", headerMark, headerLevel)); err != nil {
			return fmt.Errorf("failed to set header mark: %w", err)
		}
	}
	return nil
}

// h1H2H3 is a helper function used by the H1, H2, and H3 functions.
func (e *Excel) h1H2H3(title string, level int) error {
	cell, err := excelize.CoordinatesToCellName(e.Col, e.Row)
	if err != nil {
		return err
	}
	if err := e.f.SetCellStr(e.sheet, cell, title); err != nil {
		return err
	}
	if err := e.MarkHeader(level); err != nil {
		return err
	}

	return nil
}

// H1 creates a level 1 header and sets the specified title in the cell.
// Before setting the header, CR() executes.
// After setting the header, LF() executes.
//
// Example:
//
//	e.H1("これはレベル1のヘッダ")
func (e *Excel) H1(title string) error {
	e.Col = 1
	err := e.h1H2H3(title, 1)
	e.Row++
	return err
}

// H2 creates a level 2 header and sets the specified title in the cell.
// Before setting the header, CR().LF(2) executes.
// After setting the header, LF() executes.
//
// Example:
//
//	e.H2("これはレベル2のヘッダ")
func (e *Excel) H2(title string) error {
	e.Col, e.Row = 1, e.Row+2
	err := e.h1H2H3(title, 2)
	e.Row++
	return err
}

// H3 creates a level 3 header and sets the specified title in the cell.
// Before setting the header, CR().LF(2) executes.
// After setting the header, LF() executes.
//
// Example:
//
//	e.H3("これはレベル3のヘッダ")
func (e *Excel) H3(title string) error {
	e.Col, e.Row = 1, e.Row+2
	err := e.h1H2H3(title, 3)
	e.Row++
	return err
}

// WriteCaut writes a caution message.
func (e *Excel) WriteCaut(lines []string) error {
	e.CR(2).LF()
	cell1, err := e.Cell()
	if err != nil {
		return err
	}
	for _, s := range lines {
		if err := e.LF().SetVal(s); err != nil {
			return err
		}
	}
	cell2, err := e.CR(maxRightCellNumber).Cell()
	if err != nil {
		return err
	}
	if err := e.DrawBorders2(cell1, cell2, TBorderCaution); err != nil {
		return err
	}
	return nil
}

// WriteNote writes a note message.
func (e *Excel) WriteNote(lines []string) error {
	e.CR(2).LF()
	cell1, err := e.Cell()
	if err != nil {
		return err
	}
	for _, s := range lines {
		if err := e.LF().SetVal(s); err != nil {
			return err
		}
	}
	cell2, err := e.CR(maxRightCellNumber).Cell()
	if err != nil {
		return err
	}
	if err := e.DrawBorders2(cell1, cell2, TBorderNote); err != nil {
		return err
	}
	return nil
}

// WriteNote writes a info message.
func (e *Excel) WriteInfo(lines []string) error {
	e.CR(2).LF()
	cell1, err := e.Cell()
	if err != nil {
		return err
	}
	for _, s := range lines {
		if err := e.LF().SetVal(s); err != nil {
			return err
		}
	}
	cell2, err := e.CR(maxRightCellNumber).Cell()
	if err != nil {
		return err
	}
	if err := e.DrawBorders2(cell1, cell2, TBorderInfo); err != nil {
		return err
	}
	return nil
}

// WriteDF writes DataFrame.
// Default border type is TBorderHHeader.
//
// Example:
//
//	err := e.WriteDF(df)
//	err := e.WriteDF(df, TBorderHHeaderG)
//	err := e.WriteDF(df, TBorderVHeader)
func (e *Excel) WriteDF(df *dataframe.DataFrame,
	borderType ...TomatoBorderType) error {
	if df == nil {
		return errors.New("invalid input: the provided DataFrame is nil")
	}
	bType := TBorderHHeader
	if len(borderType) > 0 {
		bType = borderType[0]
	}
	switch bType {
	case TBorderHHeader, TBorderHHeaderG, TBorderVHeader:
		// valid cases; no action needed
	default:
		return fmt.Errorf("invalid border type: %v", bType)
	}
	cell1, _ := e.CR(2).LF().Cell()
	for _, h := range df.Headers {
		c, _ := excelize.ColumnNameToNumber(h.ColumnName)
		e.CR(c).SetVal(h.Name)
	}
	for _, values := range df.Records {
		e.LF()
		for i, h := range df.Headers {
			e.CR(h.Col).SetVal(values[i])
			i++
		}
	}
	e.Col = maxRightCellNumber
	cell2, _ := e.Cell()
	if err := e.DrawBorders2(cell1, cell2, bType); err != nil {
		return err
	}

	return nil
}
