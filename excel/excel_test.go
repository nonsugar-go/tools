package excel

import (
	"fmt"
	"path/filepath"
	"sort"
	"strconv"
	"testing"
	"time"

	"github.com/nonsugar-go/tools/excel/dataframe"
	"github.com/xuri/excelize/v2"
)

func TestExcel_NewExcel(t *testing.T) {
	dir := t.TempDir()
	filename := filepath.Join(dir, "Book1.xlsx")
	tests := []struct {
		name     string
		fontSize float64
		ok       bool
	}{
		{"size=11", 11, true},
		{"size=10", 10, true},
		{"size=409", 409, true},
		{"size=10.5", 10.5, true},
		{"size=10.4", 10.4, false},
		{"size=408.3", 408.3, false},
		{"size=409.5", 409.5, false},
		{"size=0", 0, false},
		{"size=0", -1, false},
	}
	for _, tt := range tests {
		e, err := New(filename, tt.fontSize)
		if tt.ok && err != nil {
			t.Errorf("NewExcel: want no error, but %v", err)
		}
		if !tt.ok && err == nil {
			t.Errorf("NewExcel: want error, but %v", err)
		}
		if e != nil {
			if err := e.Close(); err != nil {
				t.Errorf("Close: want no error, but %v", err)
			}
		}
	}
}

func TestExcel_OpenExcel(t *testing.T) {
	dir := "./testdata"
	filename := filepath.Join(dir, "Sample1.xlsx")
	e, err := OpenExcel(filename)
	if err != nil {
		t.Fatalf("OpenExcel: want no error, but %v", err)
	}
	defer func() {
		if err := e.Close(); err != nil {
			t.Errorf("Close: want no error, but %v", err)
		}
	}()
	file := e.GetFile()
	defFont, _ := file.GetDefaultFont()
	if defFont != defaultFont {
		t.Errorf("GetDefaultfont: want %s, but %s", defaultFont, defFont)
	}
}

func TestExcel_SaveAndClose(t *testing.T) {
	dir := "./testdata"
	filename := filepath.Join(dir, "output.xlsx")
	e, err := New(filename)
	if err != nil {
		t.Fatalf("NewExcel: want no error, but %v", err)
	}
	defer func() {
		if err := e.SaveAndClose(); err != nil {
			t.Errorf("SaveAndClose: want no error, but %v", err)
		}
	}()

	// シート「表紙」
	if err := e.NewSheet("Excel モジュールのテスト",
		SheetTypeCover); err != nil {
		t.Errorf("NewSheet: want no error, but: %v", err)
	}

	// シート「目次」
	if err := e.NewSheet("目次", SheetTypeTOC); err != nil {
		t.Errorf("NewSheet: want no error, but: %v", err)
	}

	// シート「標準」
	if err := e.NewSheet("「標準」のシートの例", SheetTypeNormal); err != nil {
		t.Errorf("NewSheet: want no error, but: %v", err)
	}
	if err := e.LF().H1("これはレベル1のヘッダ"); err != nil {
		t.Errorf("H1: want no error, but: %v", err)
	}
	if err := e.LF().H2("これはレベル2のヘッダ"); err != nil {
		t.Errorf("H1: want no error, but: %v", err)
	}
	if err := e.LF().H3("これはレベル3のヘッダ"); err != nil {
		t.Errorf("H1: want no error, but: %v", err)
	}
	if err := e.LF().SetVal("レベル1のヘッダ"); err != nil {
		t.Errorf("SetVal: want no error, but: %v", err)
	}
	if err := e.MarkHeader(1); err != nil {
		t.Errorf("MarkHeader: want no error, but: %v", err)
	}
	if err := e.LF(2).SetVal("レベル2のヘッダ"); err != nil {
		t.Errorf("SetVal: want no error, but: %v", err)
	}
	if err := e.MarkHeader(2); err != nil {
		t.Errorf("MarkHeader: want no error, but: %v", err)
	}
	e.CR(2).LF(2)
	if err := e.SetRow(&[]any{"数", 12, 13, 14, 15, 16}); err != nil {
		t.Errorf("SetRow: want no error, but: %v", err)
	}
	if err := e.LF().SetRow(&[]any{"このセルには文章を入力します。"}); err != nil {
		t.Errorf("SetRow: want no error, but: %v", err)
	}
	if err := e.CR(5).LF().SetStyle(
		NewStyle().Bold().add(b3T, b3L, b1B),
	); err != nil {
		t.Errorf("SetStyle: want no error, but: %v", err)
	}
	for y := 10; y < 20; y++ {
		for x := 3; x < 16; x++ {
			cell, err := excelize.CoordinatesToCellName(x, y)
			if err != nil {
				t.Errorf("want no error, but: %v", err)
			}
			if y == 10 {
				if err = e.SetStyleForCell(cell, NewStyle(b3T)); err != nil {
					t.Errorf("want no error, but: %v", err)
				}
			}
			if y == 19 {
				if err = e.SetStyleForCell(cell, NewStyle(b3B)); err != nil {
					t.Errorf("want no error, but: %v", err)
				}
			}
			if x == 3 {
				if err = e.SetStyleForCell(cell, NewStyle(b3L)); err != nil {
					t.Errorf("want no error, but: %v", err)
				}
			}
			if x == 15 {
				if err = e.SetStyleForCell(cell, NewStyle(b3R)); err != nil {
					t.Errorf("want no error, but: %v", err)
				}
			}
		}
	}
	if err := e.CR(2).LF().AddComment("コメントの例です"); err != nil {
		t.Errorf("AddComment: want no error, but: %v", err)
	}
	if err := e.SetVal("定数 CellStyle"); err != nil {
		t.Errorf("SetVal: want no error, but: %v", err)
	}
	e.CR(2).LF()
	for _, v := range []any{
		styleNormal,
		fontBold,
		b1T, b1L, b1R, b1B,
		b2T, b2L, b2R, b2B,
		b3T, b3L, b3R, b3B,
		bdT, bdL, bdR, bdB,
	} {
		if err := e.LF().SetRow(&[]any{v}); err != nil {
			t.Errorf("SetRow: want no error, but: %v", err)
		}
	}
	e.CR().LF(3)
	if err := e.SetVal("CellStyle map"); err != nil {
		t.Errorf("SetVal: want no error, but: %v", err)
	}
	if err := e.MarkHeader(3); err != nil {
		t.Errorf("MarkHeader: want no error, but: %v", err)
	}
	e.CR(2).LF()
	if err := e.LF().SetRow(&[]any{
		"cellStyleIDs", nil, nil, nil, nil, e.cellStyleIDs}); err != nil {
		t.Errorf("SetRow: want no error, but: %v", err)
	}
	if err := e.LF().SetRow(&[]any{
		"cellStyleMap", nil, nil, nil, nil, e.cellStyleMap}); err != nil {
		t.Errorf("SetRow: want no error, but: %v", err)
	}
	e.CR().LF(3)
	if err := e.SetVal("背景の例"); err != nil {
		t.Errorf("SetVal: want no error, but: %v", err)
	}
	if err := e.MarkHeader(1); err != nil {
		t.Errorf("MarkHeader: want no error, but: %v", err)
	}
	e.CR(2).LF()
	for _, bg := range []struct {
		name  string
		style cellStyle
	}{
		{"fg濃い青", fontHyperLink},
		{"bg灰色1", fillGray1},
		{"bg灰色2", fillGray2},
		{"bg灰色3", fillGray3},
		{"bg黄色", fillYellow},
		{"bg薄い青", fillLightBlue},
	} {
		if err := e.LF().SetVal(bg.name); err != nil {
			t.Errorf("SetVal: want no error, but: %v", err)
		}
		if err := e.SetStyle(
			NewStyle(bg.style),
		); err != nil {
			t.Errorf("SetStyle: want no error, but: %v", err)
		}
	}
	e.CR().LF(3)
	if err := e.SetVal("コメントの一覧"); err != nil {
		t.Errorf("SetVal: want no error, but: %v", err)
	}
	if err := e.MarkHeader(1); err != nil {
		t.Errorf("MarkHeader: want no error, but: %v", err)
	}
	e.CR(2).LF()
	comments, err := e.GetSortedComments(e.sheet)
	if err != nil {
		t.Errorf("GetSortedComments: want no error, but: %v", err)
	}
	for _, comment := range comments {
		if err := e.LF().SetRow(&[]any{
			comment.Cell, nil, nil, nil, comment.Paragraph,
		}); err != nil {
			t.Errorf("SetVal: want no error, but: %v", err)
		}
	}
	// シート「罫線の例」
	_ = e.NewSheet("罫線の例", SheetTypeNormal)
	_ = e.H2("各種罫線の例")

	if err := e.WriteDF((*dataframe.DataFrame)(nil)); err == nil {
		t.Errorf("WriteDF: want error, but: %v", err)
	}
	if err := e.WriteDF(
		dataframe.New("B", "ID"), TBorderSchedule); err == nil {
		t.Errorf("WriteDF: want error, but: %v", err)
	}

	_ = e.H3("セルの書式設定: デフォルト値")
	df := dataframe.New(
		"B", "タブ", "E", "大項目", "I", "小項目", "Q", "値", "U", "備考").
		Add("配置", "文字の配置", "横位置", "標準", "").
		Add("配置", "文字の配置", "縦位置", "中央揃え", "").
		Add("配置", "文字の制御", "折り返して全体を表示する", "false", "").
		Add("配置", "文字の制御", "縮小して全体を表示する", "false", "").
		Add("配置", "文字の制御", "セルを結合する", "false", "")
	if err := e.WriteDF(df); err != nil {
		t.Errorf("WriteDF: want no error, but: %v", err)
	}

	_ = e.H3("セルの書式設定: 水平ヘッダのある表")
	df = dataframe.New(
		"B", "タブ", "E", "大項目", "I", "小項目", "Q", "値", "U", "備考").
		Add("配置", "文字の配置", "横位置", "標準", "ヘッダ部分は「中央揃え」").
		Add("配置", "文字の配置", "縦位置", "中央揃え", "").
		Add("配置", "文字の制御", "折り返して全体を表示する", "false", "").
		Add("配置", "文字の制御", "縮小して全体を表示する", "true", "").
		Add("配置", "文字の制御", "セルを結合する", "true", "")
	if err := e.WriteDF(df, TBorderHHeader); err != nil {
		t.Errorf("WriteDF: want no error, but: %v", err)
	}

	_ = e.H3("セルの書式設定: 垂直ヘッダのある表")
	df = dataframe.New(
		"B", "タブ", "E", "大項目", "I", "小項目", "Q", "値", "U", "備考").
		Add("配置", "文字の配置", "横位置", "標準", "ヘッダ部分は「中央揃え」").
		Add("配置", "文字の配置", "縦位置", "中央揃え", "").
		Add("配置", "文字の制御", "折り返して全体を表示する", "false", "").
		Add("配置", "文字の制御", "縮小して全体を表示する", "true", "").
		Add("配置", "文字の制御", "セルを結合する", "true", "")
	if err := e.WriteDF(df, TBorderVHeader); err != nil {
		t.Errorf("WriteDF: want no error, but: %v", err)
	}

	_ = e.H3("セルの書式設定: 水平ヘッダのある表 (グループ対応)")
	df = dataframe.New(
		"B", "タブ", "E", "大項目", "I", "小項目", "Q", "値", "U", "備考").
		Add("配置", "文字の配置", "横位置", "標準", "ヘッダ部分は「中央揃え」").
		Add("", "", "縦位置", "中央揃え", "-").
		Add("配置", "文字の制御", "折り返して全体を表示する", "false", "-").
		Add("", "", "縮小して全体を表示する", "true", "-").
		Add("", "", "セルを結合する", "true", "-")
	if err := e.WriteDF(df, TBorderHHeaderG); err != nil {
		t.Errorf("WriteDF: want no error, but: %v", err)
	}

	_ = e.H3("セルの書式設定: 入れ子構造の表")
	cell1, _ := e.CR(2).LF().Cell()
	_ = e.SetVal("タブ")
	_ = e.CR(17).SetVal("配置")
	_ = e.CR(3).LF().SetVal("文字の配置")
	_ = e.CR(4).LF().SetVal("横位置")
	_ = e.CR(17).SetVal("標準")
	_ = e.CR(4).LF().SetVal("縦位置")
	_ = e.CR(17).SetVal("中央揃え")
	_ = e.CR(3).LF().SetVal("文字の制御")
	_ = e.CR(4).LF().SetVal("折り返して全体を表示する")
	_ = e.CR(17).SetVal("false")
	_ = e.CR(4).LF().SetVal("縮小して全体を表示する")
	_ = e.CR(17).SetVal("true")
	_ = e.CR(4).LF().SetVal("セルを結合する")
	_ = e.CR(17).SetVal("true")
	e.Col = maxRightCellNumber
	cell2, _ := e.Cell()
	if err := e.DrawBorders2(cell1, cell2, TBorderNested); err != nil {
		t.Errorf("DrawBorders2: want no error, but: %v", err)
	}

	_ = e.H2("注意")

	_ = e.H3("警告の例")
	if err := e.WriteCaut([]string{
		"警告の例です。",
		"特に気を付ける必要があることを記載します。"}); err != nil {
		t.Errorf("WriteCaut: want no error, but: %v", err)
	}

	_ = e.H3("注意の例")
	if err := e.WriteNote([]string{
		"注意の例です。",
		"注意すべき事項を記載します。"}); err != nil {
		t.Errorf("WriteCaut: want no error, but: %v", err)
	}

	_ = e.H3("ヒント (TIPS) の例")
	if err := e.WriteInfo([]string{
		"ヒント (TIPS) の例です。",
		"情報提供の目的で記載します。"}); err != nil {
		t.Errorf("WriteCaut: want no error, but: %v", err)
	}

	// シート「スタイル」
	if err := e.NewSheet("スタイル", SheetTypeNormal); err != nil {
		t.Errorf("NewSheet: want no error, but: %v", err)
	}
	if err := e.LF().H2("罫線の例"); err != nil {
		t.Errorf("H2: want no error, but: %v", err)
	}
	borderTests := []struct {
		name, cell1, cell2 string
		typ                BorderType
		isErr              bool
		fill               cellStyle
	}{
		// error
		{"X5:X5", "X5", "X5", BorderContinuousWeight1, true, fillGray3},
		// boxes
		{"C7:E9 weight=1", "C7", "E9", BorderContinuousWeight1, false, fillGray3},
		{"G7:I9 weight=7", "G7", "I9", BorderContinuousWeight2, false, fillGray3},
		{"K7:M9 weight=7", "K7", "M9", BorderContinuousWeight3, false, fillGray3},
		{"O7:Q9 double", "O7", "Q9", BorderDoubleWeight3, false, fillGray3},
		{"S7:U9 dash", "S7", "U9", BorderDashWeight1, false, fillGray3},
		// horizontal
		{"C11:E11 weight=1", "C11", "E11", BorderContinuousWeight1, false, fillGray3},
		{"G11:I11 weight=11", "G11", "I11", BorderContinuousWeight2, false, fillGray3},
		{"K11:M11 weight=11", "K11", "M11", BorderContinuousWeight3, false, fillGray3},
		{"O11:Q11 double", "O11", "Q11", BorderDoubleWeight3, false, fillGray3},
		{"S11:U11 dash", "S11", "U11", BorderDashWeight1, false, fillGray3},
		// vertical
		{"C13:C16 weight=1", "C13", "C16", BorderContinuousWeight1, false, fillGray3},
		{"G13:G16 weight=13", "G13", "G16", BorderContinuousWeight2, false, fillGray3},
		{"K13:K16 weight=13", "K13", "K16", BorderContinuousWeight3, false, fillGray3},
		{"O13:O16 double", "O13", "O16", BorderDoubleWeight3, false, fillGray3},
		{"S13:U16 dash", "S13", "U16", BorderDashWeight1, false, fillGray3},
		// nested boxes
		{"C20:Z30 weight=1", "C20", "Z30", BorderContinuousWeight1, false, fillGray1},
		{"D21:Y29 weight=1", "D21", "Y29", BorderContinuousWeight2, false, fillGray2},
		{"E22:X28 weight=13", "E22", "X28", BorderContinuousWeight3, false, fillGray3},
		{"F23:W27 double", "F23", "W27", BorderDoubleWeight3, false, fillGray3},
		{"G24:V26 dash", "G24", "V26", BorderDashWeight1, false, fillGray4},

		{"G23:V27 weight=1", "G23", "V27", BorderContinuousWeight1, false, fillRed},
		{"H22:U28 weight=1", "H22", "U28", BorderContinuousWeight2, false, fillYellow},
		{"I21:T29 weight=13", "I21", "T29", BorderContinuousWeight3, false, fillLightBlue},
		{"J20:S30 double", "J20", "S30", BorderDoubleWeight3, false, fillPurple},
		{"K19:R31 dash", "K19", "R31", BorderDashWeight1, false, fillOrange},
	}
	for _, tt := range borderTests {
		err := e.DrawBorders(tt.cell1, tt.cell2, tt.typ)
		if err != nil && !tt.isErr {
			t.Errorf("DrawBorders: want error, but: %v", err)
		}
		if err == nil && tt.isErr {
			t.Errorf("DrawBorders: want no error, but: %v", err)
		}
	}
	for _, tt := range borderTests {
		if !tt.isErr {
			if err := e.SetStyleForCellRange(
				tt.cell1, tt.cell2, NewStyle().add(tt.fill)); err != nil {
				t.Errorf("SetStyleForCellRange: want no error, but: %v", err)
			}
		}
	}

	e.Row = 32
	if err := e.LF().H2("フォントサイズの例"); err != nil {
		t.Errorf("H2: want no error, but: %v", err)
	}
	if err := e.CR(2).LF().SetVal("フォントサイズ20"); err != nil {
		t.Errorf("SetVal: want no error, but: %v", err)
	}
	if err := e.SetStyle(NewStyle().add(fontSize20)); err != nil {
		t.Errorf("SetStyle: want no error, but: %v", err)
	}

	if err := e.CR(2).LF().SetVal("フォントサイズ20"); err != nil {
		t.Errorf("SetVal: want no error, but: %v", err)
	}
	if err := e.SetStyle(NewStyle().add(fontSize20)); err != nil {
		t.Errorf("SetStyle: want no error, but: %v", err)
	}
	if err := e.SetVal("フォントサイズ12"); err != nil {
		t.Errorf("SetVal: want no error, but: %v", err)
	}
	if err := e.SetStyle(NewStyle().add(fontSize12)); err != nil {
		t.Errorf("SetStyle: want no error, but: %v", err)
	}

	if err := e.LF(3).H2("フォントの色"); err != nil {
		t.Errorf("H2: want no error, but: %v", err)
	}

	fontColorTests := []struct {
		name      string
		fontColor cellStyle
	}{
		{"濃い赤", fontDeepRed},
		{"赤", fontRed},
		{"オレンジ", fontOrange},
		{"黄", fontYellow},
		{"薄い緑", fontLightGreen},
		{"緑", fontGreen},
		{"薄い青", fontLightBlue},
		{"青", fontBlue},
		{"濃い青", fontDarkBlue},
		{"紫", fontPurple},
		{"ハイパーリンク用", fontHyperLink},
	}
	for _, tt := range fontColorTests {
		if err := e.CR(2).LF().SetVal(tt.name); err != nil {
			t.Errorf("SetVal: want no error, but: %v", err)
		}
		if err := e.SetStyle(NewStyle().add(tt.fontColor)); err != nil {
			t.Errorf("SetStyle: want no error, but: %v", err)
		}
	}

	if err := e.LF(3).H2("背景色"); err != nil {
		t.Errorf("H2: want no error, but: %v", err)
	}

	fillColorTests := []struct {
		name      string
		fontColor cellStyle
	}{
		{"濃い赤", fillDeepRed},
		{"赤", fillRed},
		{"オレンジ", fillOrange},
		{"黄", fillYellow},
		{"薄い緑", fillLightGreen},
		{"緑", fillGreen},
		{"薄い青", fillLightBlue},
		{"青", fillBlue},
		{"濃い青", fillDarkBlue},
		{"紫", fillPurple},
		{"グレー1", fillGray1},
		{"グレー2", fillGray2},
		{"グレー3", fillGray3},
		{"グレー4", fillGray4},
		{"グレー5", fillGray5},
		{"<-- Excel Macro のヘッダ1", fillHeaderColor1},
		{"<-- Excel Macro のヘッダ2", fillHeaderColor2},
		{"<-- Excel Macro のヘッダ3", fillHeaderColor3},
		{"ピンク <-- Excel Macro の CAUTION用", fillCaution},
		{"黄色 <-- Excel Macro のNOTE用", fillNote},
		{"薄い青 <-- Excel Macro のHINT用", fillHint},
	}
	for _, tt := range fillColorTests {
		if err := e.CR(2).LF().SetVal(tt.name); err != nil {
			t.Errorf("SetVal: want no error, but: %v", err)
		}
		if err := e.SetStyle(NewStyle().add(tt.fontColor)); err != nil {
			t.Errorf("SetStyle: want no error, but: %v", err)
		}
	}

	// シート「セルスタイルのデータ」
	_ = e.NewSheet("セルスタイルのデータ", SheetTypeNormal)
	_ = e.H2("セルスタイルのデータ")
	_ = e.H3("cellStyleIDs")
	df = dataframe.New("B", "Key", "Q", "cellStyleIDs")
	keys := make([]cellStyle, 0, len(e.cellStyleIDs))
	for k := range e.cellStyleIDs {
		keys = append(keys, k)
	}
	sort.Slice(keys, func(i, j int) bool {
		return uint64(keys[i]) < uint64(keys[j])
	})
	for _, k := range keys {
		v := e.cellStyleIDs[k]
		df.Add(strconv.FormatUint(uint64(k), 10), strconv.Itoa(v))
	}
	if err := e.WriteDF(df); err != nil {
		t.Errorf("WriteDF: want no error, but: %v", err)
	}
	_ = e.H3("cellStyleMap")
	df = dataframe.New("B", "Key", "Q", "cellStyleMap")
	keys2 := make([]string, 0, len(e.cellStyleMap))
	for k := range e.cellStyleMap {
		keys2 = append(keys2, k)
	}
	sort.Strings(keys2)
	for _, k := range keys2 {
		v := e.cellStyleMap[k]
		df.Add(k, strconv.FormatUint(uint64(v), 10))
	}
	if err := e.WriteDF(df); err != nil {
		t.Errorf("WriteDF: want no error, but: %v", err)
	}

	// シート「サンプル テスト」
	if err := e.NewSheet("サンプル テスト"); err != nil {
		t.Errorf("NewSheet: want no error, but: %v", err)
	}
	tests := []struct {
		name string
		x, y int
	}{
		{name: "(3, 3)", x: 3, y: 3},
		{name: "(3, 4)", x: 3, y: 4},
		{name: "(5, 5)", x: 5, y: 5},
	}
	for _, tt := range tests {
		e.Col, e.Row = tt.x, tt.y
		err := e.SetRow(&[]any{
			fmt.Sprintf("(%d, %d)", e.Col, e.Row), "文字列", nil, -1, 3.14})
		if err != nil {
			t.Errorf("SetRow: %s: want no error, but %v", tt.name, err)
		}
	}

	// シート「人物リスト」
	if err := e.NewSheet("人物リスト"); err != nil {
		t.Errorf("NewSheet: want no error, but: %v", err)
	}
	if err := e.SetActiveSheet(); err != nil {
		t.Errorf("NewActiveSheet: want no error, but: %v", err)
	}
	if err := e.SetHeader([]Header{
		{"No", 0},
		{"姓", 7},
		{"名", 7},
		{"人物の説明", 24},
	}); err != nil {
		t.Errorf("SetHader: want no error, but %v", err)
	}
	rows := [][]any{
		{1, "大谷", "翔平", "ベースボール プレイヤー"},
		{2, "鈴木", "一郎", "殿堂入り"},
		{3, "野茂", "英雄", "トルネード投法"},
	}
	for _, row := range rows {
		if err := e.LF().SetRow(&row); err != nil {
			t.Errorf("SetRow: want no error, but %v", err)
		}
	}
	if err := e.AddTable("Table1"); err != nil {
		t.Errorf("AddTable: want no error, but %v", err)
	}
}

func Test_ColumnNumberToName(t *testing.T) {
	tests := []struct {
		name   string
		colNum int
		expect string
		isErr  bool
	}{
		{"col number is 5", 5, "E", false},
		{"col number is 6", 6, "F", false},
		{"col number is 0", 0, "", true},
		{"col number is -1", -1, "", true},
	}

	for _, tt := range tests {
		got, err := ColumnNumberToName(tt.colNum)
		if !tt.isErr && err != nil {
			t.Errorf("%s: want no error, but %v", tt.name, err)
		}
		if got != tt.expect {
			t.Errorf("%s: want %s, but %s", tt.name, tt.expect, got)
		}
	}
}

func Test_CordinatesToCellName(t *testing.T) {
	tests := []struct {
		name     string
		col, row int
		abs      bool
		expect   string
		isErr    bool
	}{
		{"col: 1, row: 2", 1, 2, false, "A2", false},
		{"col: 3, row: 1", 3, 1, true, "$C$1", false},
		{"col: 0, row: 1", 0, 1, false, "", true},
	}

	for _, tt := range tests {
		got, err := CoordinatesToCellName(tt.col, tt.row, tt.abs)
		if !tt.isErr && err != nil {
			t.Errorf("%s: want no error, but %v", tt.name, err)
		}
		if got != tt.expect {
			t.Errorf("%s: want %s, but %s", tt.name, tt.expect, got)
		}
	}
}

func TestExcel_Cell(t *testing.T) {
	tests := []struct {
		name     string
		col, row int
		expect   string
		isErr    bool
	}{
		{"col: 1, row: 2", 1, 2, "A2", false},
		{"col: 3, row: 1", 3, 1, "C1", false},
		{"col: 0, row: 1", 0, 1, "", true},
		{"col: 10, row: 99999999999999", 0, 1, "", true},
	}

	e, _ := New("dummy.xlsx")
	defer e.Close()
	for _, tt := range tests {
		e.Col, e.Row = tt.col, tt.row
		got, err := e.Cell()
		if !tt.isErr && err != nil {
			t.Errorf("%s: want no error, but %v", tt.name, err)
		}
		if got != tt.expect {
			t.Errorf("%s: want %s, but %s", tt.name, tt.expect, got)
		}
	}
}

func TestExcelMakeHeader(t *testing.T) {
	tests := []struct {
		name  string
		level int
		isErr bool
	}{
		{"level=-1", -1, true},
		{"level=0", 0, false},
		{"level=1", 1, false},
		{"level=3", 3, false},
		{"level=4", 4, true},
	}

	e, _ := New("dummy.xlsx")
	defer e.Close()
	_ = e.NewSheet("foo")
	for _, tt := range tests {
		err := e.MarkHeader(tt.level)
		if !tt.isErr && err != nil {
			t.Errorf("%s: want no error, but %v", tt.name, err)
		}
		if tt.isErr && err == nil {
			t.Errorf("%s: want error, but %v", tt.name, err)
		}
	}
}

func TestExcel_GetLastColumnNumberAndGetLastRowNumber(t *testing.T) {
	tests := []struct {
		name     string
		sheet    string
		col, row int
	}{
		{"test1 1 5", "test 1", 1, 5},
		{"test2 3 8", "test 2", 3, 8},
		{"test3 9 12", "test 3", 9, 12},
	}

	e, _ := New("dummy.xlsx")
	defer e.Close()
	for _, tt := range tests {
		_ = e.NewSheet(tt.name)
		_ = e.SetVal(tt.name, tt.col, tt.row)
		lastC, err := e.GetLastColumnNumber()
		if err != nil {
			t.Errorf("%s: GetLastColumnNumber: want no error, but %v",
				tt.name, err)
		}
		lastR, err := e.GetLastRowNumber()
		if err != nil {
			t.Errorf("%s: GetLastRowNumber: want no error, but %v", tt.name, err)
		}
		if lastC != tt.col {
			t.Errorf("%s: GetLastColumnNumber: want %d, but %d",
				tt.name, tt.col, lastC)
		}
		if lastR != tt.row {
			t.Errorf("%s: GetLastColumnNumber: want %d, but %d",
				tt.name, tt.row, lastR)
		}
	}

	tt := tests[0]
	lastC, err := e.GetLastColumnNumber(tt.name)
	if err != nil {
		t.Errorf("%s: GetLastColumnNumber: want no error, but %v",
			tt.name, err)
	}
	lastR, err := e.GetLastRowNumber(tt.name)
	if err != nil {
		t.Errorf("%s: GetLastRowNumber: want no error, but %v", tt.name, err)
	}
	if lastC != tt.col {
		t.Errorf("%s: GetLastColumnNumber: want %d, but %d",
			tt.name, tt.col, lastC)
	}
	if lastR != tt.row {
		t.Errorf("%s: GetLastColumnNumber: want %d, but %d",
			tt.name, tt.row, lastR)
	}
}

func TestRelCellNameToAbsCellName(t *testing.T) {
	tests := []struct {
		name     string
		cell     string
		expected string
	}{
		{"A1", "A1", "$A$1"},
		{"BZ1234", "BZ1234", "$BZ$1234"},
		{"$H$20", "$H$20", "$H$20"},
	}

	e, _ := New("dummy.xlsx")
	defer e.Close()
	_ = e.NewSheet("dummy")
	for _, tt := range tests {
		got, err := RelCellNameToAbsCellName(tt.cell)
		if err != nil {
			t.Errorf("%s: want no error, but %v",
				tt.name, err)
		}
		if got != tt.expected {
			t.Errorf("%s: want %s, but %s",
				tt.name, tt.expected, got)
		}
	}
}

func TestExcel_GetVal(t *testing.T) {
	tests := []struct {
		name     string
		col, row int
		value    any
		expected string
	}{
		{"A1 Hello", 1, 1, "Hello", "Hello"},
		{"B2 \"123\"", 2, 2, "123", "123"},
		{"B3 ture", 2, 3, true, "TRUE"},
		{"B4 false", 2, 4, false, "FALSE"},
		{"C1 \"\"", 3, 1, "", ""},
		{"C2 nil", 3, 2, nil, ""},
	}

	e, _ := New("dummy.xlsx")
	defer e.Close()
	_ = e.NewSheet("dummy")
	for _, tt := range tests {
		_ = e.SetVal(tt.value, tt.col, tt.row)
		got, err := e.GetVal(tt.col, tt.row)
		if err != nil {
			t.Errorf("%s: want no error, but %v",
				tt.name, err)
		}
		if got != tt.expected {
			t.Errorf("%s: want %s, but %s",
				tt.name, tt.expected, got)
		}
	}
}

func TestExcel_HasVal(t *testing.T) {
	tests := []struct {
		name     string
		col, row int
		value    any
		expected bool
	}{
		{"A1 Hello", 1, 1, "Hello", true},
		{"B2 \"123\"", 2, 2, "123", true},
		{"B3 ture", 2, 3, true, true},
		{"B4 false", 2, 4, false, true},
		{"B5 time.Now()", 2, 5, time.Now(), true},
		{"C1 \"\"", 3, 1, "", false},
		{"C2 nil", 3, 2, nil, false},
	}

	e, _ := New("dummy.xlsx")
	defer e.Close()
	_ = e.NewSheet("dummy")
	for _, tt := range tests {
		_ = e.SetVal(tt.value, tt.col, tt.row)
		got, err := e.HasVal(tt.col, tt.row)
		if err != nil {
			t.Errorf("%s: want no error, but %v",
				tt.name, err)
		}
		if got != tt.expected {
			t.Errorf("%s: want %t, but %t",
				tt.name, tt.expected, got)
		}
	}
}
