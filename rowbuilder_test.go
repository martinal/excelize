package excelize

import (
	"testing"
)

func TestDenseColumns(t *testing.T) {
	cols := DenseColumns(4)
	if 4 != len(cols) {
		t.Fatalf("Invalid length")
	}
	if "D" != cols[3] {
		t.Fatalf("Invalid letter")
	}
}
