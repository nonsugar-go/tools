package main

import (
	"flag"
	"fmt"
	"log"
	"os"
	"strings"
	"text/tabwriter"

	"github.com/nonsugar-go/tools/excel"
)

func main() {
	var (
		filename                       string
		sheetIdx                       int
		colMin, colMax, rowMin, rowMax int
	)
	flag.StringVar(&filename, "in", "", "Excel ファイル (*.xlsx)")
	flag.IntVar(&sheetIdx, "sheet", 0, "sheet index")
	flag.IntVar(&colMin, "col-min", 1, "column min number")
	flag.IntVar(&colMax, "col-max", 8, "column max number")
	flag.IntVar(&rowMin, "row-min", 1, "row min number")
	flag.IntVar(&rowMax, "row-max", 8, "row max number")
	flag.Parse()
	if filename == "" {
		log.Print("Excel ファイルが指定されていません")
		os.Exit(1)
	}
	e, err := excel.OpenExcel(filename)
	if err != nil {
		log.Printf("Excel ファイルが開けません: %v", err)
		os.Exit(1)
	}
	defer func() {
		if err := e.Close(); err != nil {
			log.Print(err)
		}
	}()
	f := e.GetFile()
	var s string
	fmt.Println(strings.Repeat("-", 72))
	w := tabwriter.NewWriter(os.Stdout, 0, 0, 2, ' ', 0)
	s, _ = f.GetDefaultFont()
	fmt.Println("[WORKBOOK PROPS]")
	fmt.Fprint(w, "KEY\tVALUE\t\n")
	fmt.Fprintf(w, "Default Font\t%s\t\n", s)
	wbProps, _ := f.GetWorkbookProps()
	fmt.Fprintf(w, "CodeName\t%#v\t\n", *wbProps.CodeName)
	fmt.Fprintf(w, "Date1094\t%#v\t\n", *wbProps.Date1904)
	fmt.Fprintf(w, "FilterPrivacy\t%#v\t\n", *wbProps.FilterPrivacy)
	w.Flush()
	fmt.Println(strings.Repeat("-", 72))
	fmt.Println("[DEFINED NAME]")
	fmt.Fprint(w, "IDX\tNAME\tComment\tRefersTo\tScope\t\n")
	for i, definedName := range f.GetDefinedName() {
		fmt.Fprintf(w, "%d\t%#v\t%#v\t%#v\t%#v\t\n",
			i,
			definedName.Name,
			definedName.Comment,
			definedName.RefersTo,
			definedName.Scope,
		)
	}
	w.Flush()
	fmt.Println(strings.Repeat("-", 72))
	fmt.Println("[SHEET LIST]")
	fmt.Fprint(w, "IDX\tSHEET NAME\t\n")
	for _, sheet := range f.GetSheetList() {
		i, err := f.GetSheetIndex(sheet)
		if err != nil {
			log.Println(err)
		}
		fmt.Fprintf(w, "%d\t%s\t\n", i, sheet)
	}
	w.Flush()
	fmt.Println(strings.Repeat("-", 72))
	sheet := f.GetSheetName(sheetIdx)
	fmt.Printf("[SHEET PROPS:%s]\n", sheet)
	fmt.Fprint(w, "KEY\tVALUE\t\n")
	shProps, _ := f.GetSheetProps(sheet)
	fmt.Fprintf(w, "AutoPageBreak\t%#v\t\n", *shProps.AutoPageBreaks)
	fmt.Fprintf(w, "BaseColWidth\t%#v\t\n", *shProps.BaseColWidth)
	fmt.Fprintf(w, "CodeName\t%#v\t\n", shProps.CodeName)
	fmt.Fprintf(w, "CustomHeight\t%#v\t\n", *shProps.CustomHeight)
	fmt.Fprintf(w, "DefaultColWidth\t%#v\t\n", *shProps.DefaultColWidth)
	fmt.Fprintf(w, "DefaultRowHeight\t%#v\t\n", *shProps.DefaultRowHeight)
	fmt.Fprintf(w, "EnableFormatConditionsCalculation\t%#v\t\n",
		*shProps.EnableFormatConditionsCalculation)
	fmt.Fprintf(w, "FitToPage\t%#v\t\n", shProps.FitToPage)
	fmt.Fprintf(w, "OutlineSummaryBelow\t%#v\t\n", *shProps.OutlineSummaryBelow)
	fmt.Fprintf(w, "OutlineSummaryRight\t%#v\t\n", shProps.OutlineSummaryRight)
	fmt.Fprintf(w, "Published\t%#v\t\n", *shProps.Published)
	fmt.Fprintf(w, "TabColorIndexed\t%#v\t\n", shProps.TabColorIndexed)
	fmt.Fprintf(w, "TabColorRGB\t%#v\t\n", shProps.TabColorRGB)
	fmt.Fprintf(w, "TabColorTheme\t%#v\t\n", shProps.TabColorTheme)
	fmt.Fprintf(w, "TabColorTint\t%#v\t\n", shProps.TabColorTint)
	fmt.Fprintf(w, "ThickBottom\t%#v\t\n", *shProps.ThickBottom)
	fmt.Fprintf(w, "ThickTop\t%#v\t\n", *shProps.ThickTop)
	fmt.Fprintf(w, "ZeroHeight\t%#v\t\n", *shProps.ZeroHeight)
	w.Flush()
	fmt.Println(strings.Repeat("-", 72))
	fmt.Printf("[PAGE LAYOUT:%s]\n", sheet)
	pageLayout, _ := f.GetPageLayout(sheet)
	fmt.Fprintf(w, "AdjustTo\t%#v\t\n", *pageLayout.AdjustTo)
	fmt.Fprintf(w, "BlackAndWhite\t%#v\t\n", *pageLayout.BlackAndWhite)
	fmt.Fprintf(w, "FirstPageNumber\t%#v\t\n", *pageLayout.FirstPageNumber)
	fmt.Fprintf(w, "FitToHeight\t%#v\t\n", pageLayout.FitToHeight)
	fmt.Fprintf(w, "FitToWidth\t%#v\t\n", pageLayout.FitToWidth)
	fmt.Fprintf(w, "Orientation\t%#v\t\n", *pageLayout.Orientation)
	fmt.Fprintf(w, "Size\t%#v\t\n", *pageLayout.Size)
	w.Flush()
	fmt.Println(strings.Repeat("-", 72))
	fmt.Printf("[PAGE MARGINS:%s]\n", sheet)
	pageMargins, _ := f.GetPageMargins(sheet)
	fmt.Fprintf(w, "Bottom\t%#v\t\n", *pageMargins.Bottom)
	fmt.Fprintf(w, "Footer\t%#v\t\n", *pageMargins.Footer)
	fmt.Fprintf(w, "Header\t%#v\t\n", *pageMargins.Header)
	fmt.Fprintf(w, "Horizontally\t%#v\t\n", pageMargins.Horizontally)
	fmt.Fprintf(w, "Left\t%#v\t\n", *pageMargins.Left)
	fmt.Fprintf(w, "Right\t%#v\t\n", *pageMargins.Right)
	fmt.Fprintf(w, "Top\t%#v\t\n", *pageMargins.Top)
	fmt.Fprintf(w, "Vertically\t%#v\t\n", pageMargins.Vertically)
	w.Flush()
	fmt.Println(strings.Repeat("-", 72))
	fmt.Printf("[HEADER FOOTER:%s]\n", sheet)
	headerFooter, _ := f.GetHeaderFooter(sheet)
	fmt.Fprintf(w, "AlignWithMargins\t%#v\t\n", headerFooter.AlignWithMargins)
	fmt.Fprintf(w, "DifferentFirst\t%#v\t\n", headerFooter.DifferentFirst)
	fmt.Fprintf(w, "DifferentOddEven\t%#v\t\n", headerFooter.DifferentOddEven)
	fmt.Fprintf(w, "EvenFooter\t%#v\t\n", headerFooter.EvenFooter)
	fmt.Fprintf(w, "EvenHeader\t%#v\t\n", headerFooter.EvenHeader)
	fmt.Fprintf(w, "FirstFooter\t%#v\t\n", headerFooter.FirstFooter)
	fmt.Fprintf(w, "FirstHeader\t%#v\t\n", headerFooter.FirstHeader)
	fmt.Fprintf(w, "OddFooter\t%#v\t\n", headerFooter.OddFooter)
	fmt.Fprintf(w, "OddHeader\t%#v\t\n", headerFooter.OddHeader)
	fmt.Fprintf(w, "ScaleWithDoc\t%#v\t\n", headerFooter.ScaleWithDoc)
	w.Flush()
	fmt.Println(strings.Repeat("-", 72))
	fmt.Printf("[COLUMNS:%s]\n", sheet)
	fmt.Fprint(w, "COL\tVISIBLE\tWIDTH\t\n")
	for col := colMin; col <= colMax; col++ {
		colName, _ := excel.ColumnNumberToName(col)
		colVisible, _ := f.GetColVisible(sheet, colName)
		colWidth, _ := f.GetColWidth(sheet, colName)
		fmt.Fprintf(w, "%s\t%#v\t%#v\t\n", colName, colVisible, colWidth)
	}
	w.Flush()
	fmt.Println(strings.Repeat("-", 72))
	fmt.Printf("[ROWS:%s]\n", sheet)
	fmt.Fprint(w, "ROW\tVISIBLE\tHEIGHT\t\n")
	for row := rowMin; row <= rowMax; row++ {
		rowVisible, _ := f.GetRowVisible(sheet, row)
		rowHeight, _ := f.GetRowHeight(sheet, row)
		fmt.Fprintf(w, "%d\t%#v\t%#v\t\n", row, rowVisible, rowHeight)
	}
	w.Flush()
	fmt.Println(strings.Repeat("-", 72))
}
