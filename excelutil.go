// Package excelutil is a command line program that converts an excel workbook with potentially multiple spread sheets
// of a given format to another format while doing the appropriate maths. It can create graphs and sort the
// columns of the primary output according to the maximum value per output.
// author: Daniel Schuette (email: d.schuette@online.de)
// license: MIT license (see github.com/DanielSchuette)
package excelutil

import (
	"fmt"
	"log"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
)

// define constants
const (
	ENUM  = 1 // enumerator = 340
	DENOM = 2 // denominator = 380
	SKIP  = 3 // we don't want this field
	// background values for 340/380 are always to the last two values
)

// ExcelWorkbook holds all important workbook-related information
type ExcelWorkbook struct {
	XLSX       *excelize.File
	SheetNames []string
	NumSheets  int
	Dims       [2]int
}

// NumberOfSheets returns the number of sheets in an excelWorkbook
func (wb *ExcelWorkbook) NumberOfSheets() int {
	return len(wb.SheetNames)
}

// StartRow returns the row index at which the actual data matrix starts as an integer
func (wb *ExcelWorkbook) StartRow(sheet, label string) (int, error) {
	m := wb.XLSX.GetRows(sheet)
	for idx, val := range m {
		if string(val[0]) == label {
			return idx, nil
		}
	}
	return 0, fmt.Errorf("did not find a row with label %s in column 1", label)
}

// Dimensions returns the dimensions of a sheet in the format (rows, cols)
func (wb *ExcelWorkbook) Dimensions(sheet string) [2]int {
	m := wb.XLSX.GetRows(sheet)
	d := [2]int{
		len(m),    // size of row dimension
		len(m[0]), // size of column dimension
	}
	return d
}

// Open opens a .xlsx file and assigns it to an ExcelWorkbook
func (wb *ExcelWorkbook) Open(name string) {
	xlsx, err := excelize.OpenFile(name)
	if err != nil {
		log.Fatalf("error while opening file: %s\n", err)
	}
	wb.XLSX = xlsx
}

// GetSheetNames gets all sheet names from a given workbook and stores them in the ExcelWorkbook struct
func (wb *ExcelWorkbook) GetSheetNames() {
	sn := make([]string, 0)
	for _, n := range wb.XLSX.GetSheetMap() {
		sn = append(sn, n)
	}
	wb.SheetNames = sn
	wb.NumSheets = wb.NumberOfSheets()
}

// prints a useful delimiter
func PrintDelim() {
	for i := 0; i < 70; i++ {
		fmt.Printf("%s", "-")
		time.Sleep(10 * time.Millisecond)
	}
	fmt.Println()
}

// takes an integer and returns an Excel-style string representation of it (e.g. 1 = A, 3 = C, 27 = AA, ...)
// the current implementation only works for a limited amount of cells, though
func GetColumn(num int) string {
	num-- // because of go's 0 indexing
	alphabet := [26]string{
		"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M",
		"N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
	}
	switch {
	// return a single letter
	case num < (1 * len(alphabet)):
		return fmt.Sprintf("%s", alphabet[num])

	// return a combination of letters, starting with "A..."
	case (num >= (1 * len(alphabet))) && (num < (2 * len(alphabet))):
		return fmt.Sprintf("%s%s", "A", alphabet[num-len(alphabet)])

	// return a combination of letters, starting with "B..."
	case (num >= (2 * len(alphabet))) && (num < (3 * len(alphabet))):
		return fmt.Sprintf("%s%s", "B", alphabet[num-(2*len(alphabet))])

	// return a combination of letter
	case (num >= (3 * len(alphabet))) && (num < (4 * len(alphabet))):
		return fmt.Sprintf("%s%s", "C", alphabet[num-(3*len(alphabet))])

	// return a combination of letter
	case (num >= (4 * len(alphabet))) && (num < (5 * len(alphabet))):
		return fmt.Sprintf("%s%s", "D", alphabet[num-(4*len(alphabet))])

	// return a combination of letter
	case (num >= (5 * len(alphabet))) && (num < (6 * len(alphabet))):
		return fmt.Sprintf("%s%s", "E", alphabet[num-(5*len(alphabet))])

	// return a combination of letter
	case (num >= (6 * len(alphabet))) && (num < (7 * len(alphabet))):
		return fmt.Sprintf("%s%s", "F", alphabet[num-(6*len(alphabet))])

	// return a combination of letter
	case (num >= (7 * len(alphabet))) && (num < (8 * len(alphabet))):
		return fmt.Sprintf("%s%s", "G", alphabet[num-(7*len(alphabet))])

	// return a combination of letter
	case (num >= (8 * len(alphabet))) && (num < (9 * len(alphabet))):
		return fmt.Sprintf("%s%s", "H", alphabet[num-(8*len(alphabet))])

	// return a combination of letter
	case (num >= (9 * len(alphabet))) && (num < (10 * len(alphabet))):
		return fmt.Sprintf("%s%s", "I", alphabet[num-(9*len(alphabet))])

	// return a combination of letter
	case (num >= (10 * len(alphabet))) && (num < (11 * len(alphabet))):
		return fmt.Sprintf("%s%s", "J", alphabet[num-(10*len(alphabet))])

	// log a fatal error if none of these cases holds true
	default:
		log.Fatal("algorithm cannot work with so many input columns")
		return ""
	}
}

// FindMaxElem is a helper function for iterating over a map;
// it finds the max value ==> gets its index ==> returns the index of the max value
func FindMaxElem(input map[int]float64) int {
	maxVal := 0.0
	index := 0
	for idx, val := range input {
		if val > maxVal {
			maxVal = val
			index = idx
		}
	}
	return index
}
