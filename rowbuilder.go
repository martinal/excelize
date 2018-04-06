package excelize

import (
	"fmt"
	"strconv"
	"time"
)

func DenseColumns(n int) []string {
	if n > 26 {
		panic("Double letters not implemented")
	}
	cols := make([]string, 0, n)
	for i := 0; i < n; i++ {
		cols = append(cols, string(byte(65+i)))
	}
	return cols
}

type RowBuilder interface {
	AddRow(values []interface{}, styles []int) error
}

type rowBuilderImpl struct {
	xfile              *File
	sheet              string
	xlsx               *xlsxWorksheet
	startRow           int
	nextRow            int
	numColumns         int
	denseColumnHeaders []string
}

func NewRowBuilder(xfile *File, sheet string, startRow int, numColumns int) RowBuilder {
	return &rowBuilderImpl{
		xfile:              xfile,
		sheet:              sheet,
		xlsx:               xfile.workSheetReader(sheet),
		startRow:           startRow,
		nextRow:            startRow,
		numColumns:         numColumns,
		denseColumnHeaders: DenseColumns(numColumns),
	}
}

func convertValue(val interface{}) (value string, typ string, err error) {

	// Default to numeric type
	typ = "n"

	switch v := val.(type) {
	case nil:
		value = ""
	case string:
		value = v
		typ = "str"
	case []byte:
		value = string(v)
		typ = "str"

	case float32:
		value = strconv.FormatFloat(float64(v), 'f', -1, 32)
	case float64:
		value = strconv.FormatFloat(v, 'f', -1, 64)

	case int:
		value = strconv.FormatInt(int64(v), 10)
	case int64:
		value = strconv.FormatInt(int64(v), 10)
	case int32:
		value = strconv.FormatInt(int64(v), 10)
	case int16:
		value = strconv.FormatInt(int64(v), 10)
	case int8:
		value = strconv.FormatInt(int64(v), 10)

	case uint:
		value = strconv.FormatUint(uint64(v), 10)
	case uint64:
		value = strconv.FormatUint(uint64(v), 10)
	case uint32:
		value = strconv.FormatUint(uint64(v), 10)
	case uint16:
		value = strconv.FormatUint(uint64(v), 10)
	case uint8:
		value = strconv.FormatUint(uint64(v), 10)

	// TODO: Which value should bool have in excel?
	case bool:
		if v {
			value = "true"
		} else {
			value = "false"
		}
		typ = "b"

	// TODO: This formatting taken from excelize looks funky
	case time.Duration:
		value = strconv.FormatFloat(float64(v.Seconds()/86400), 'f', -1, 32)
		// f.setDefaultTimeStyle(sheet, axis, 22) // TODO: User nust set style separately
	case time.Time:
		value = strconv.FormatFloat(float64(timeToExcelTime(timeToUTCTime(v))), 'f', -1, 64)
		// f.setDefaultTimeStyle(sheet, axis, 22) // TODO: User nust set style separately

	case fmt.Stringer:
		value = v.String()
	default:
		value = fmt.Sprintf("%v", v)
	}

	return
}

func (b *rowBuilderImpl) AddRow(values []interface{}, styles []int) error {
	// Ensure next row exists with full column width
	// TODO: This can be inlined, simplified, and optimized using b.denseColumnHeaders
	completeRow(b.xlsx, b.nextRow, b.numColumns)

	// Get row pointer
	row := &b.xlsx.SheetData.Row[b.nextRow-1]
	b.nextRow++

	// Set styles
	for i, style := range styles {
		row.C[i].S = style
	}

	// Add values and types
	for i, value := range values {
		value, typ, err := convertValue(value)
		if err != nil {
			return err
		}
		row.C[i].V = value
		row.C[i].T = typ
	}

	return nil
}
