# tools/excel

ブックのデフォルト フォント サイズを設定するために、
`excelize` に SetDefaultFontSize() を追加しています。

## vendor フォルダー

```bash
go mod vendor // OR go work vendor
cd vendor/github.com/xuri/excelize/v2
cp styles.go{,.orig}
vi styles.go
```

```
diff styles.go.orig styles.go
1795c1795
< func (f *File) SetDefaultFont(fontName string) error {
---
> func (f *File) SetDefaultFont(fontName string, fontSize ...float64) error {
1800a1801,1803
>       if len(fontSize) > 0 {
>               font.Sz.Val = float64Ptr(fontSize[0])
>       }
```

