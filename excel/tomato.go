package excel

import (
	"errors"
	"fmt"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

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
		e.Col, e.Row = 1, 4
	case SheetTypeTOC:
		e.Col, e.Row = 10, 5
		if err := e.MakeTOC(); err != nil {
			return err
		}
		e.Col, e.Row = 1, 1
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
	case TBorderNested:
		// TODO:
		// ' 入れ子構造の表
		// If drawLevel < 0 Or drawLevel > 9 Then Exit Sub
		// paramBorders (drawLevel)
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
		panic("invalid TomatoBorderType")
	}

	return nil
}

func (e *Excel) headerBorders(borderType TomatoBorderType,
	col1, row1, col2, row2 int) error {
	isWrapText := 2
	// 0: 変更しない
	// 1: 折り返して全体を表示する
	// 2: 縮小して全体を表示する

	// 区切りとなるセル位置を記録する
	separateCol := make(map[int]struct{})
	separateRow := make(map[int]struct{})
	// separateRw(UBound(separateRw)) = 1 // TODO

	for r := row1; r <= row2; r++ {
		nMergedColumns := 1
		for c, prevC := col1, col1; c <= col2; c++ {
			cell, err := excelize.CoordinatesToCellName(c, r)
			if err != nil {
				return err
			}
			if r == row1 {
				// 1行目 かつ 何か値がある 場合
				// 区切りとなるセルの位置を覚える
				value, err := e.f.GetCellValue(e.sheet, cell)
				if err != nil {
					return err
				}
				if value != "" {
					separateCol[c] = struct{}{}
				}
			} // if r == row1
			if c == col1 {
				// 1列目 かつ 何か値がある 場合
				// 区切りとなるセルの位置を覚える
				value, err := e.f.GetCellValue(e.sheet, cell)
				if err != nil {
					return err
				}
				if value != "" {
					separateRow[r] = struct{}{}
				}
			} else { // nColumns != 1
				if _, ok := separateCol[c]; ok {
					// 1列目でなく かつ 区切りとなるセル の場合
					// セルの結合
					cell1, err := excelize.CoordinatesToCellName(prevC, r)
					if err != nil {
						return err
					}
					cell2, err := excelize.CoordinatesToCellName(c-1, r)
					if err != nil {
						return err
					}
					if err := e.headerBorders_mergeCells(cell1, cell2,
						borderType, r-row1+1, nMergedColumns, isWrapText); err != nil {
						return err
					}
					nMergedColumns++
					prevC = c
				} // if _, ok := separateCol[c]; ok
				if c == col2 {
					// 最終列だった場合
					// セルの結合
					cell1, err := excelize.CoordinatesToCellName(prevC, r)
					if err != nil {
						return err
					}
					cell2, err := excelize.CoordinatesToCellName(c, r)
					if err != nil {
						return err
					}
					if err := e.f.MergeCell(e.sheet, cell1, cell2); err != nil {
						return err
					}
					if err := e.headerBorders_mergeCells(cell1, cell2,
						borderType, r-row1+1, nMergedColumns, isWrapText); err != nil {
						return err
					}
				} // if c = col2
			} // if c == col1
		} // for c
	} // for r
	/*
				    If hOrV = "G" Then
				        ' グループ対応だったら、複数の値が無いグループを結合
				        Application.DisplayAlerts = False

				        nColumns = 0

				        ' 各列でループする。
				        For Each cl In Selection.Columns
				            nColumns = nColumns + 1
				            nRows = 0

				            If separateCl(nColumns) <> 0 Then
				                ' 列の区切りの場合

				                ' 各行でループする。
				                For Each rw In cl.Rows
				                    nRows = nRows + 1

				                    If nRows = 1 Then
				                        Set prevSeparateCl = rw.Offset(1, 0)
				                        numElements = 0
				                    Else
				                        If rw.Value <> "" Then
				                            numElements = numElements + 1
				                        End If

				                        If separateRw(nRows + 1) <> 0 Then
				                            ' 行区切りか最終行 (番兵) の場合
				                            If numElements < 2 Then
				                                ' セルの結合
				                                mergeCells2 Range(prevSeparateCl, rw), isWrapText, False
				                                numElements = 0
				                            End If

				                            Set prevSeparateCl = rw.Offset(1, 0)
				                            numElements = 0
				                        End If
				                    End If
				                Next rw
				            End If
				        Next cl
				    End If

				    ' 外枠を太線で引く
				    outsideBorders (Selection.cells) // done
		////////////////////////////
		Dim rw As Range ' 行
		Dim cl As Range ' 列
		Dim separateCl() As Integer ' 区切りとなる列方向のセル
		Dim separateRw() As Integer ' 区切りとなる行方向のセル
		Dim prevSeparateCl As Range ' ひとつ前の区切りとなるセル
		Dim nRows As Single ' 処理中の行目
		Dim nColumns As Single ' 処理中の列目
		Dim nMergedColumns As Single ' 処理中のマージ後の列目
		Dim numElements ' 項目数を数える
		Dim isWrapText As Single ' 1.折り返し, 2.縮小, -1.変更無し
	*/

	// 外枠を太線で引く
	//outsideBorders(Selection.cells)
	cell1, err := excelize.CoordinatesToCellName(col1, row1)
	if err != nil {
		return err
	}
	cell2, err := excelize.CoordinatesToCellName(col2, row2)
	if err != nil {
		return err
	}
	if err := e.DrawBorders(cell1, cell2,
		BorderContinuousWeight2); err != nil {
		return err
	}

	return nil
}

func (e *Excel) headerBorders_mergeCells(
	topLeftCell, bottomRightCell string, // cells
	borderType TomatoBorderType, nRows int, nMergedColumns int, isWrapText int,
) error {
	//// hOrV "H", "G", "V" -> borderType

	// セルの結合
	//////////////// TODO -->
	// mergeCells2 cells, isWrapText
	if err := e.f.MergeCell(e.sheet,
		topLeftCell, bottomRightCell); err != nil {
		return err
	}

	switch borderType {
	case TBorderHHeader, TBorderHHeaderG:
		// 水平の場合
	case TBorderVHeader:
		// 垂直の場合
	default:
		panic("invalid TomatoBorderType")
	}
	//////////////// TODO <--
	/*
		    If hOrV = "H" Or hOrV = "G" Then
		        ' 水平の場合

		        ' 右に縦線を実線で引く
		        With cells.Borders(xlEdgeRight)
		            .lineStyle = xlContinuous
		            .weight = xlThin
		        End With

		        If nRows = 1 Then
		            ' 1行目の場合
		            ' 下に横線を2重線で引く
		            With cells.Borders(xlEdgeBottom)
		                .lineStyle = xlDouble
		                .weight = xlThick
		            End With

		            ' ヘッダに色を塗る
		            With cells.Interior
		                .ColorIndex = HEADER_COLOR
		                .Pattern = xlSolid
		                .PatternColorIndex = xlAutomatic
		            End With

		            ' 横位置: 中央揃え
		            cells.HorizontalAlignment = xlCenter
		        Else
		            ' 1行目以外の場合
		            ' 下に横線を実線で引く
		            With cells.Borders(xlEdgeBottom)
		                .lineStyle = xlContinuous
		                If hOrV = "G" And ActiveSheet.cells(cells.Row + 1, 2) = "" Then
		                    ' 次の行の一列目が空欄なら、下に横線を破線を引く
		                    .lineStyle = xlDash
		                End If
		                .weight = xlThin
		            End With
		        End If
		    Else
		        ' 垂直の場合

		        ' 下に横線を実線で引く
		        With cells.Borders(xlEdgeBottom)
		            .lineStyle = xlContinuous
		            .weight = xlThin
		        End With

		        If nMergedColumns = 1 Then
		            ' マージ後の1列目の場合
		            ' 右に縦線を2重線で引く
		            With cells.Borders(xlEdgeRight)
		                .lineStyle = xlDouble
		                .weight = xlThick
		            End With

		            ' ヘッダに色を塗る
		            With cells.Interior
		                .ColorIndex = HEADER_COLOR
		                .Pattern = xlSolid
		                .PatternColorIndex = xlAutomatic
		            End With

		            ' 横位置: 中央揃え
		            cells.HorizontalAlignment = xlCenter
		        Else
		            ' マージ後の1列目以外の場合
		            ' 右に縦線を実線で引く
		            With cells.Borders(xlEdgeRight)
		                .lineStyle = xlContinuous
		                .weight = xlThin
		            End With
		        End If
		    End If
		    Exit Sub
		End Sub
	*/
	return nil
}

func (e *Excel) admonitionBorders(borderType TomatoBorderType,
	col1, row1, col2, row2 int) error {
	// TODO
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
	e.Col = 1
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
	e.Row++

	return nil
}

// H1 creates a level 1 header and sets the specified title in the cell.
// Before setting the header, CR() executes.
// After setting the header, LF() executes.
//
// Example:
//
//	LF().H1("これはレベル1のヘッダ")
//	LF().H2("これはレベル2のヘッダ")
//	LF().H3("これはレベル3のヘッダ")
func (e *Excel) H1(title string) error {
	return e.h1H2H3(title, 1)
}

// H2 creates a level 2 header and sets the specified title in the cell.
func (e *Excel) H2(title string) error {
	return e.h1H2H3(title, 2)
}

// H3 creates a level 3 header and sets the specified title in the cell.
func (e *Excel) H3(title string) error {
	return e.h1H2H3(title, 3)
}
