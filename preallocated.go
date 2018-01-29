package excelize

import (
	"encoding/xml"
	"fmt"
	"strconv"
)

func (f *File) GetSheet(sheet string) *xlsxWorksheet {
	return f.workSheetReader(sheet)
}

// SetCellIntPreallocated provides function to set int type value of a cell by given
// worksheet name, cell coordinates and cell value.
func SetCellIntPreallocated(xlsx *xlsxWorksheet, xAxis, yAxis int, value int) {
	xlsx.SheetData.Row[xAxis].C[yAxis].S = prepareCellStyle(xlsx, yAxis+1, xlsx.SheetData.Row[xAxis].C[yAxis].S)
	xlsx.SheetData.Row[xAxis].C[yAxis].T = ""
	xlsx.SheetData.Row[xAxis].C[yAxis].V = strconv.Itoa(value)
}

// SetCellStrPreallocated provides function to set string type value of a cell. Total number
// of characters that a cell can contain 32767 characters.
func SetCellStrPreallocated(xlsx *xlsxWorksheet, xAxis, yAxis int, value string) {

	if xAxis >= len(xlsx.SheetData.Row) {
		panic(fmt.Sprintf("Invalid row index %d >= %d", xAxis, len(xlsx.SheetData.Row)))
	}
	if yAxis >= len(xlsx.SheetData.Row[xAxis].C) {
		panic(fmt.Sprintf("Invalid col index %d >= %d", yAxis, len(xlsx.SheetData.Row[xAxis].C)))
	}

	if len(value) > 32767 {
		value = value[0:32767]
	}
	// Leading space(s) character detection.
	if len(value) > 0 {
		if value[0] == 32 {
			xlsx.SheetData.Row[xAxis].C[yAxis].XMLSpace = xml.Attr{
				Name:  xml.Name{Space: NameSpaceXML, Local: "space"},
				Value: "preserve",
			}
		}
	}

	xlsx.SheetData.Row[xAxis].C[yAxis].S = prepareCellStyle(xlsx, yAxis+1, xlsx.SheetData.Row[xAxis].C[yAxis].S)
	xlsx.SheetData.Row[xAxis].C[yAxis].T = "str"
	xlsx.SheetData.Row[xAxis].C[yAxis].V = value
}

// SetCellDefaultPreallocated provides function to set string type value of a cell as
// default format without escaping the cell.
func SetCellDefaultPreallocated(xlsx *xlsxWorksheet, xAxis, yAxis int, value string) {
	xlsx.SheetData.Row[xAxis].C[yAxis].S = prepareCellStyle(xlsx, yAxis+1, xlsx.SheetData.Row[xAxis].C[yAxis].S)
	xlsx.SheetData.Row[xAxis].C[yAxis].T = ""
	xlsx.SheetData.Row[xAxis].C[yAxis].V = value
}

// SetCellValue provides function to set value of a cell. The following shows
// the supported data types:
//
//    int
//    int8
//    int16
//    int32
//    int64
//    uint
//    uint8
//    uint16
//    uint32
//    uint64
//    float32
//    float64
//    string
//    []byte
//    time.Time
//    nil
//
// Note that default date format is m/d/yy h:mm of time.Time type value. You can
// set numbers format by SetCellStyle() method.
// func SetCellValuePreallocated(xlsx *xlsxWorksheet, xAxis, yAxis int, value interface{}) {
// 	switch t := value.(type) {
// 	case int:
// 		SetCellInt(xlsx, xAxis, yAxis, value.(int))
// 	case int8:
// 		SetCellInt(xlsx, xAxis, yAxis, int(value.(int8)))
// 	case int16:
// 		SetCellInt(xlsx, xAxis, yAxis, int(value.(int16)))
// 	case int32:
// 		SetCellInt(xlsx, xAxis, yAxis, int(value.(int32)))
// 	case int64:
// 		SetCellInt(xlsx, xAxis, yAxis, int(value.(int64)))
// 	case uint:
// 		SetCellInt(xlsx, xAxis, yAxis, int(value.(uint)))
// 	case uint8:
// 		SetCellInt(xlsx, xAxis, yAxis, int(value.(uint8)))
// 	case uint16:
// 		SetCellInt(xlsx, xAxis, yAxis, int(value.(uint16)))
// 	case uint32:
// 		SetCellInt(xlsx, xAxis, yAxis, int(value.(uint32)))
// 	case uint64:
// 		SetCellInt(xlsx, xAxis, yAxis, int(value.(uint64)))
// 	case float32:
// 		SetCellDefault(xlsx, xAxis, yAxis, strconv.FormatFloat(float64(value.(float32)), 'f', -1, 32))
// 	case float64:
// 		SetCellDefault(xlsx, xAxis, yAxis, strconv.FormatFloat(float64(value.(float64)), 'f', -1, 64))
// 	case string:
// 		SetCellStr(xlsx, xAxis, yAxis, t)
// 	case []byte:
// 		SetCellStr(xlsx, xAxis, yAxis, string(t))
// 	case time.Time:
// 		SetCellDefault(xlsx, xAxis, yAxis, strconv.FormatFloat(float64(timeToExcelTime(timeToUTCTime(value.(time.Time)))), 'f', -1, 64))
// 	case nil:
// 		SetCellStr(xlsx, xAxis, yAxis, "")
// 	default:
// 		SetCellStr(xlsx, xAxis, yAxis, fmt.Sprintf("%v", value))
// 	}
// }

// AlphaStrings creates a slice of strings alphaStrings such that alphaStrings[j] == ToAlphaString(j) for j=0,...,n-1
func AlphaStrings(n int) []string {
	alphaStrings := make([]string, n)
	for j := 0; j < n; j++ {
		alphaStrings[j] = ToAlphaString(j)
	}
	return alphaStrings
}

// EnsureSheetSize ensures that a given sheet has the requested number of rows and cols allocated
func (f *File) EnsureSheetSize(sheet string, rows, cols int) {
	xlsx := f.workSheetReader(sheet)

	for i := len(xlsx.SheetData.Row); i < rows; i++ {
		xlsx.SheetData.Row = append(xlsx.SheetData.Row, xlsxRow{R: i + 1})
	}

	alphaStrings := AlphaStrings(cols)

	for i := range xlsx.SheetData.Row {
		istr := strconv.Itoa(i + 1)
		for j := len(xlsx.SheetData.Row[i].C); j < cols; j++ {
			xlsx.SheetData.Row[i].C = append(xlsx.SheetData.Row[i].C, xlsxC{R: alphaStrings[j] + istr})
		}
	}

	m := len(xlsx.SheetData.Row)
	n := len(xlsx.SheetData.Row[m-1].C)
	fmt.Printf("Allocated: %d  %d\n", m, n)
}
