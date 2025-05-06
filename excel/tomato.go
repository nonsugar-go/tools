package excel

import (
	"fmt"
	"strconv"
	"strings"
)

/*

' ////////////////////////////////////////////////////////////////
'  コメントにヘッダの印を付けるか目次を作成する。
' ////////////////////////////////////////////////////////////////
Public Sub markHeaderOrMakeTOC()
    On Error GoTo ERR1
    Application.ScreenUpdating = False

    ' 変数定義
    Dim msgResult As String ' 質問の結果
    Dim headerLevel As Integer ' ヘッダのレベル

    ' レベルいくつのヘッダの印を付けますか？
    msgResult = InputBox("レベルいくつのヘッダの印を付けますか？[1-" & MAX_HEADER_LEVEL & ", 0, T]" _
        & vbCrLf & "1-" & MAX_HEADER_LEVEL & " ... ヘッダの印を付ける (レベル)" _
        & vbCrLf & "0 ... コメントを削除する" _
        & vbCrLf & "T ... 目次を作成する" _
        , , 2)
    If msgResult = "" Then Exit Sub
    If InStr(1, "TtＴｔ", Left$(msgResult, 1), vbTextCompare) <> 0 Then
        ' 目次を作成する
        makeTOC
        Exit Sub
    ElseIf Not IsNumeric(msgResult) Then
        Exit Sub
    Else
        headerLevel = CInt(msgResult)
        If headerLevel >= 0 And headerLevel <= MAX_HEADER_LEVEL Then
            markHeader (headerLevel)
        End If
    End If
    Exit Sub
ERR1: errorExit (Err.Description)
End Sub


'  コメントにヘッダの印を付ける
Private Sub markHeader(ByVal headerLevel As Integer)
    On Error GoTo ERR1

    If headerLevel = 0 Then
        ' コメントを削除する
        Selection.ClearComments
    Else
        ' ヘッダの印を付ける
        With Selection
            .Font.Bold = True
            .ClearComments
            .AddComment
            .comment.Visible = False
            .comment.Text Text:=HEADER_MARK & headerLevel
        End With
    End If
    Exit Sub
ERR1: errorExit (Err.Description)
End Sub
*/

/*
' 目次を作る
Private Sub makeTOC(Optional isInteractive As Boolean = True)
    On Error GoTo ERR1

    ' 変数定義
    Dim Sh As Worksheet ' 処理中のシート
    Dim number(3) As Integer ' 各レベルの番号
    Dim i As Integer ' ループ変数
    Dim cl As Range ' 列
    Dim endY As Long ' 最終行
    Dim y As Long ' ループ変数 (行)
    Dim strTmp As String ' 一時変数
    Dim headerLevel As Integer ' ヘッダのレベル
    Dim msgResult As Variant ' MsgBox の結果
    Dim isTOC As Boolean ' 目次を作成
    Dim objComment As Object ' コメント
    Dim TocCell As Range ' TOC を挿入する位置のセル
    Dim headers(1000) As String ' ヘッダの配列
    Dim headerLevels(1000) As Integer ' ヘッダのレベルの配列
    Dim cellOfHeaders(1000) As String ' ヘッダが位置するセルを示す文字列
    Dim NumOfHeaders As Integer ' ヘッダの数
    Dim findBeginCell As Range ' 目次開始のコメントがあるセル
    Dim findEndCell As Range ' 目次終了のコメントがあるセル
    Dim insertCellRow As Long ' 挿入位置の行番号
    Dim insertCellColumn As Long ' 挿入位置の列番号
    Dim tmpName As name ' 名前

    If isInteractive Then
    msgResult = MsgBox("目次を作成します。" & vbCrLf _
        & vbCrLf & "[はい] ... この位置または以前目次を作成した場所に目次を挿入" _
        & vbCrLf & "[いいえ] ... 目次は作成せず、ヘッダのナンバリングの更新のみ" _
        & vbCrLf & "[キャンセル] ... 処理の中止" _
        , vbQuestion Or vbYesNoCancel)
    Else
        msgResult = vbYes
    End If
    If msgResult = vbYes Then
        isTOC = True
    ElseIf msgResult = vbNo Then
        isTOC = False
    Else
        Exit Sub
    End If

    Set TocCell = ActiveCell
    For i = 0 To 3
        number(i) = 0
    Next i
    NumOfHeaders = 0

    ' 名前の定義を削除
    For Each tmpName In ActiveWorkbook.Names
        If Right$(tmpName.name, Len(SUF_HEADER_NAME)) = SUF_HEADER_NAME Then tmpName.Delete
    Next

    For Each Sh In Worksheets
        endY = endRow(Sh)
        For y = 1 To endY
            Set cl = Sh.cells(y, 1)
            While (cl.Column < MAX_EXCEL_COLUMN)
                Set objComment = cl.comment
                If Not objComment Is Nothing Then
                    If Left$(objComment.Text, Len(HEADER_MARK)) = HEADER_MARK Then
                        strTmp = Mid$(cl.comment.Text, Len(HEADER_MARK) + 1, 1)
                        If IsNumeric(strTmp) Then
                            headerLevel = CInt(strTmp)
                            If headerLevel >= 1 And headerLevel <= MAX_HEADER_LEVEL Then
                                ' セルがヘッダの場合
                                number(headerLevel) = number(headerLevel) + 1
                                strTmp = ""
                                For i = 1 To headerLevel
                                    strTmp = strTmp & number(i) & "."
                                Next i
                                For i = headerLevel + 1 To MAX_HEADER_LEVEL
                                    number(i) = 0
                                Next i

                                For i = 1 To Len(cl.Value)
                                    If InStr("0123456789.０１２３４５６７８９．", Mid$(cl.Value, i, 1)) = 0 Then
                                        strTmp = strTmp & " " & Trim(Mid$(cl.Value, i))
                                        i = Len(cl.Value) ' For i ループを抜ける
                                    End If
                                Next i
                                cl.Value = strTmp
                                headers(NumOfHeaders) = strTmp

                                ' 名前を定義する。
                                strTmp = StrConv(strTmp, vbNarrow)

                                strTmp = Replace(strTmp, " ", "_")
                                strTmp = Replace(strTmp, "　", "_")
                                strTmp = Replace(strTmp, "!", "_")
                                strTmp = Replace(strTmp, Chr$(34), "_") ' "
                                strTmp = Replace(strTmp, "#", "_")
                                strTmp = Replace(strTmp, "$", "_")
                                strTmp = Replace(strTmp, "%", "_")
                                strTmp = Replace(strTmp, "&", "_")
                                strTmp = Replace(strTmp, "'", "_")
                                strTmp = Replace(strTmp, "(", "_")
                                strTmp = Replace(strTmp, ")", "_")
                                strTmp = Replace(strTmp, "*", "_")
                                strTmp = Replace(strTmp, "+", "_")
                                strTmp = Replace(strTmp, ",", "_")
                                strTmp = Replace(strTmp, "-", "_")
                                strTmp = Replace(strTmp, "/", "_")
                                strTmp = Replace(strTmp, ":", "_")
                                strTmp = Replace(strTmp, ";", "_")
                                strTmp = Replace(strTmp, "<", "_")
                                strTmp = Replace(strTmp, "=", "_")
                                strTmp = Replace(strTmp, ">", "_")
                                strTmp = Replace(strTmp, "@", "_")
                                strTmp = Replace(strTmp, "[", "_")
                                strTmp = Replace(strTmp, "]", "_")
                                strTmp = Replace(strTmp, "^", "_")
                                strTmp = Replace(strTmp, "`", "_")
                                strTmp = Replace(strTmp, "{", "_")
                                strTmp = Replace(strTmp, "|", "_")
                                strTmp = Replace(strTmp, "}", "_")
                                strTmp = Replace(strTmp, "~", "_")

                                strTmp = PRE_HEADER_NAME & strTmp & SUF_HEADER_NAME
                                ' ActiveWorkbook.Names.Add Name:=strTmp, RefersTo:=cl

                                headerLevels(NumOfHeaders) = headerLevel
                                cellOfHeaders(NumOfHeaders) = "'" & Sh.name & "'!" & cl.Address(True, True, xlA1)
                                NumOfHeaders = NumOfHeaders + 1
                            End If
                        End If
                    End If
                End If
                Set cl = cl.End(xlToRight)
            Wend
        Next y
    Next Sh

    ' 目次を作成する
    Set findBeginCell = findComment(BEGIN_TABLE_OF_CONTENTS)
    If isTOC Then
        If Not findBeginCell Is Nothing Then
            ' 目次開始のコメントが見つかった場合
            Set findEndCell = findEndComment(findBeginCell, END_TABLE_OF_CONTENTS)
            If Not findEndCell Is Nothing Then
                ' コメントがある位置に挿入する。
                findBeginCell.Worksheet.Activate
                findBeginCell.Activate
                insertCellRow = findBeginCell.Row
                insertCellColumn = findBeginCell.Column
                Range(findBeginCell, findEndCell).EntireRow.Select
                Selection.Delete Shift:=xlUp
                ActiveSheet.cells(insertCellRow, insertCellColumn).Activate
            End If
        Else
            ' 目次開始のコメントが見つからなかった場合
            ' カーソルのあったセルの位置に目次を挿入する｡
            TocCell.Activate
        End If

        For i = 0 To NumOfHeaders - 1
            If i <> 0 And headerLevels(i) = 1 Then
                ' レベル1 のヘッダの前で空行を挿入
                ActiveCell.EntireRow.Insert
                ActiveCell.Offset(1, 0).Activate
            End If

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
    Exit Sub
ERR1: errorExit (Err.Description)
End Sub


' あたえられた文字列で始まるコメントが最初に現れるセルを探す。
' ただし、コメントが記入されたセルに値が入っている必要がある。
Private Function findComment(ByVal comment As String) As Range
    On Error GoTo ERR1

    ' 変数定義
    Dim Sh As Worksheet ' 処理中のシート
    Dim endY As Long ' 最終行
    Dim cl As Range ' 列
    Dim y As Long ' ループ変数 (行)
    Dim objComment As Object ' コメント

    For Each Sh In Worksheets
        endY = endRow(Sh)
        For y = 1 To endY
            Set cl = Sh.cells(y, 1)
            While (cl.Column < MAX_EXCEL_COLUMN)
                Set objComment = cl.comment
                If Not objComment Is Nothing Then
                    If Left$(objComment.Text, Len(comment)) = comment Then
                        ' コメントが見つかった場合
                        Set findComment = cl
                        Exit Function
                    End If
                End If
                Set cl = cl.End(xlToRight)
            Wend
        Next y
    Next Sh

    Set findComment = Nothing
    Exit Function
ERR1: errorExit (Err.Description)
End Function


' あたえられたセルと同じ列で、下方向に最初に現れるあたえられた文字列で始まるコメントが付いているセルを探す。
' コメントが記入されたセルに値が入っている必要はない。
' 但し、そのシートの一番下に値の入っている行 + 100行 以下は見ない。
Private Function findEndComment(ByVal beginCell As Range, ByVal comment As String) As Range
    On Error GoTo ERR1

    ' 変数定義
    Dim Sh As Worksheet ' シート
    Dim endY As Long ' 最終行
    Dim cl As Range ' 列
    Dim y As Long ' ループ変数 (行)
    Dim objComment As Object ' コメント

    Set Sh = beginCell.Worksheet
    endY = endRow(Sh) + 100
    If endY > MAX_EXCEL_ROW Then endY = MAX_EXCEL_ROW
    For y = beginCell.Row + 1 To endY
        Set cl = Sh.cells(y, beginCell.Column)
        Set objComment = cl.comment
        If Not objComment Is Nothing Then
            If Left$(objComment.Text, Len(comment)) = comment Then
                ' コメントが見つかった場合
                Set findEndComment = cl
                Exit Function
            End If
        End If
    Next y

    Set findEndComment = Nothing
    Exit Function
ERR1: errorExit (Err.Description)
End Function
*/

// func endRow --> GetLastRowNumber()

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
		if i == 0 {
			// 最初のヘッダのみ､開始コメントを付ける｡
			if err := e.AddComment(beginTableOfContents); err != nil {
				return fmt.Errorf("failed to add comment in cell '%s': %w",
					cell, err)
			}
		}
	}
	// 最終行の場合のみ、終了コメントを付ける。
	if err := e.AddComment(endTableOfContents); err != nil {
		return fmt.Errorf("failed to add comment in cell '%s': %w",
			cell, err)
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
		if err := e.SetCellStyleForCurrentCell(NewStyle().Bold()); err != nil {
			return fmt.Errorf("failed to set header mark: %w", err)
		}
		if err := e.AddComment(
			fmt.Sprintf("%s%d", headerMark, headerLevel)); err != nil {
			return fmt.Errorf("failed to set header mark: %w", err)
		}
	}
	return nil
}
