package excelize

import "strconv"

// AddRowValues adds a new row at end of sheet with provided column values
func AddRowValues(f *File, sheet string, values map[string]interface{}) error {

	xlsx := f.workSheetReader(sheet)

	// Find next row
	srow := strconv.Itoa(len(xlsx.SheetData.Row) + 1)

	// Insert all values (can be optimized)
	for scol, val := range values {
		f.SetCellValue(sheet, scol+srow, val)
	}

	return nil
}
