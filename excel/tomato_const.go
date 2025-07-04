package excel

// tomato 関連定数
const (
	progName     = "図書作成支援マクロ 「TOMATO (TOsyo MAcro TO be happy)」"
	progDesc     = "図書作成を支援してくれる MS-Excel マクロです。"
	progFilename = "TOMATO-0319.XLSB"
	progVersion  = "Version 3.19"
	// ttFontSize  = 10 // コード用(TeleType)のフォントサイズ
	// ttRowHeight = 13.5 // コード用(TeleType)の行間
	// textRightCell = "AF" → "AG"
	leftMargin         = 1.6
	rightMargin        = 0.6
	topMargin          = 1.6
	bottomMargin       = 1.6
	headerMargin       = 0.6
	footerMargin       = 0.6
	maxRightCell       = "AG"
	maxRightCellNumber = 33 // "AG", Excel Macro には無い
	sa4LeftMargin      = 1.6
	sa4RightMargin     = 1.6
	sa4TopMargin       = 1.6
	sa4BottomMargin    = 1.6
	sa4HeaderMargin    = 0.6
	sa4FooterMargin    = 0.6
	sa4MaxRightCell    = "DE"
	sa4MaxRightRow     = "73"
	sa3MaxRightCell    = "FD"
	sa3MaxRightRow     = "109"
	// defaultFont = "ＭＳ Ｐゴシック" // デフォルトのフォント
	// defaultFont          = "游ゴシック"   // デフォルトのフォント
	// defaultFontSize      = 10        // デフォルトのフォントサイズ
	ttFont               = "ＭＳ ゴシック" // コード用(TeleType)のフォント
	ttFontSize           = 10        // コード用(TeleType)のフォントサイズ
	ttRowHeight          = 13.5      // コード用(TeleType)の行間
	ttCont               = "⇒"       // 継続行の記号
	tabWidth             = 8         // TAB の文字数
	pageFormat           = "&P / &N" // ページのフォーマット
	headerColor          = 15        // ヘッダの塗りつぶしの色
	color1               = 192       // 濃い赤   #C00000
	color2               = 255       // 赤       #FF0000
	color3               = 49407     // オレンジ #FFC000
	color4               = 65535     // 黄       #FFFF00
	color5               = 5296274   // 薄い緑   #50D050
	color6               = 5287936   // 緑       #50B000
	color7               = 15773696  // 薄い青   #F0F0F0
	color8               = 12611584  // 青       #C0C0FF
	color9               = 6299648   // 濃い青   #6020C0
	color10              = 10498160  // 紫       #A080B0
	cautionColor         = 26        // ピンク
	noteColor            = 27        // 黄色
	hintColor            = 28        // 薄い青
	levelColor1          = 16
	levelColor2          = 48
	levelColor3          = 15
	levelColor4          = 16
	levelColor5          = 48
	levelColor6          = 15
	levelColor7          = 16
	levelColor8          = 48
	levelColor9          = 15
	defaultTextColumns   = 80
	headerMark           = "TOMATO: Header"
	maxHeaderLevel       = 3
	maxExcelRow          = 65536
	maxExcelColumn       = 256
	maxSelectionRow      = 10000
	preHeaderName        = "■"
	sufHeaderName        = "_TOMATO"
	tocName              = "■0._目次_TOMATO"
	beginTableOfContents = "TOMATO: Begin{Table of Contents}"
	endTableOfContents   = "TOMATO: End{Table of Contents}"
	beginTextFile        = "TOMATO: Begin{Text File}"
	endTextFile          = "TOMATO: End{Text File}"
	// ■□●○✓✔☑'☒
	checkBox = "\u25A0\u25A1\u25CF\u25CB\u2713\u2714\u2611\u2612"
)
