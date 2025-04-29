//go:generate stringer -type CellType -trimprefix CellType const.go
package main

// CellType is the type of cell value type.
type CellType byte

// Cell value types enumeration.
const (
	CellTypeUnset CellType = iota
	CellTypeBool
	CellTypeDate
	CellTypeError
	CellTypeFormula
	CellTypeInlineString
	CellTypeNumber
	CellTypeSharedString
)
