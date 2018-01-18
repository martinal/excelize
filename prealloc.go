package excelize

import (
	"strconv"
)

// AlphaStrings creates a slice of strings alphaStrings such that alphaStrings[j] == ToAlphaString(j) for j=0,...,n-1
func AlphaStrings(n int) []string {
	alphaStrings := make([]string, n)
	for j := 0; j < n; j++ {
		alphaStrings[j] = ToAlphaString(j)
	}
	return alphaStrings
}

// EnsureSheetSize ensures that a given sheet has the requested number of rows and cols allocated
func (f *File) EnsureSheetSize(sheetName string, rows, cols int) {
	data := &f.Sheet[sheetName].SheetData

	for i := len(data.Row); i < rows; i++ {
		data.Row = append(data.Row, xlsxRow{R: i + 1})
	}

	alphaStrings := AlphaStrings(cols)

	for i, rowdata := range data.Row {
		istr := strconv.Itoa(i + 1)
		for j := len(rowdata.C); j < cols; j++ {
			rowdata.C = append(rowdata.C, xlsxC{R: alphaStrings[j] + istr})
		}
	}
}
