# tools/excel

ブックのデフォルト フォント サイズを設定するために、
`excelize` に SetDefaultFontSize() を追加しています。

## vendor フォルダー

```bash
go mod vendor // or go work vendor
cd vendor/github.com/xuri/excelize/v2
cp styles.go{,.orig}
vi styles.go
```

```
diff styles.go.orig styles.go
1809a1810,1826
> // SetDefaultFontAndSize changes the default font in the workbook.
> func (f *File) SetDefaultFontAndSize(fontName string, fontSize float64) error {
> 	font, err := f.readDefaultFont()
> 	if err != nil {
> 		return err
> 	}
> 	font.Name.Val = stringPtr(fontName)
> 	font.Sz.Val = &fontSize
> 	f.mu.Lock()
> 	s, _ := f.stylesReader()
> 	f.mu.Unlock()
> 	s.Fonts.Font[0] = font
> 	custom := true
> 	s.CellStyles.CellStyle[0].CustomBuiltIn = &custom
> 	return err
> }
> 
```

