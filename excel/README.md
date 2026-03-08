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

```diff
diff -u styles.go.orig styles.go
--- styles.go.orig      2026-03-08 22:48:54.135527631 +0900
+++ styles.go   2026-03-08 22:52:36.312296983 +0900
@@ -1834,12 +1834,15 @@
 }

 // SetDefaultFont changes the default font in the workbook.
-func (f *File) SetDefaultFont(fontName string) error {
+func (f *File) SetDefaultFont(fontName string, fontSize ...float64) error {
        font, err := f.readDefaultFont()
        if err != nil {
                return err
        }
        font.Name.Val = stringPtr(fontName)
+       if len(fontSize) > 0 {
+               font.Sz.Val = float64Ptr(fontSize[0])
+       }
        f.mu.Lock()
        s, _ := f.stylesReader()
        f.mu.Unlock()
```
