// ExcelUtil is a command line program that converts an excel workbook with potentially multiple spread sheets
// of a given format to another format while doing the appropriate maths. It can create graphs and sort the
// columns of the primary output according to the maximum value per output.
// author: Daniel Schuette (email: d.schuette@online.de)
// license: MIT license (see github.com/DanielSchuette)

package main

import (
	"flag"
	"fmt"
	"log"
	"sort"
	"strconv"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
)

// define flags
var (
	xlsxName = flag.String("file_path", "", "specify the path to the Excel (.xlsx) file that you want to process")

	responseThreshold = flag.Float64("threshold", 1.2, "not yet implemented!\noptional argument specifying a response threshold (as a floating point number)\nevery column without a value larger than this number will be dropped during analysis\nif you don't want this behavior, override it by putting in '0'")

	trimOutput = flag.Int("trimmed_output", 450, "specify after how many measurements the output should be trimmed\nthis option applies only to the '_ratios.xlsx' output file")

	addChart = flag.Bool("add_chart", false, "--add_chart=true adds two line plots visualizing the first 12 columns of every sheet (defaults to false)\nonly the first up to 470 measurements are plotted and the plots are drawn at columns A470 and R470\nmake sure to change this hard-coded format if your experimental setup/sampling-interval changes")

	verbose = flag.Bool("verbose", false, "--verbose=true results in an (extremely) verbose output (defaults to false)")

	sortStart = flag.Int("start", 30, "specify at which measurement you want to start looking for a peak that is then used to sort columns")

	sortEnd = flag.Int("stop", 360, "specify at which measurement you want to stop looking for a peak that is then used to sort columns")

	printMap = flag.Bool("print_order", true, "--print_order=false does not print the ordered max values for all cells in all sheets to stdout")
)

// define constants
const (
	ENUM  = 1 // enumerator = 340
	DENOM = 2 // denominator = 380
	SKIP  = 3 // we don't want this field
	// background values for 340/380 are always to the last two values
)

// excelWorkbook holds all important workbook-related information
type excelWorkbook struct {
	XLSX       *excelize.File
	SheetNames []string
	NumSheets  int
	Dims       [2]int
}

// numSheets returns the number of sheets in an excelWorkbook
func (wb *excelWorkbook) numSheets() int {
	return len(wb.SheetNames)
}

// startRow returns the row index at which the actual data matrix starts as an integer
func (wb *excelWorkbook) startRow(sheet, label string) (int, error) {
	m := wb.XLSX.GetRows(sheet)
	for idx, val := range m {
		if string(val[0]) == label {
			return idx, nil
		}
	}
	return 0, fmt.Errorf("did not find a row with label %s in column 1", label)
}

// dims returns the dimensions of a sheet in the format (rows, cols)
func (wb *excelWorkbook) dims(sheet string) [2]int {
	m := wb.XLSX.GetRows(sheet)
	d := [2]int{
		len(m),    // size of row dimension
		len(m[0]), // size of column dimension
	}
	return d
}

func main() {
	// defer done statement
	defer printDelim()
	defer fmt.Println("done")

	// parse flags and check for errors
	printDelim()
	flag.Parse()
	if *xlsxName == "" {
		log.Fatal("provide a correct file path (see --help)")
	}

	// start to process data
	fmt.Printf("opened file: %s\n", *xlsxName)
	fmt.Println("starting to process data...")

	// create a new ExcelWorkbook
	wb := &excelWorkbook{}

	// open .xlsx file
	xlsx, err := excelize.OpenFile(*xlsxName)
	if err != nil {
		log.Fatalf("error while opening file: %s\n", err)
	}
	wb.XLSX = xlsx

	// get sheet names and store in slice
	sn := make([]string, 0)
	for _, n := range wb.XLSX.GetSheetMap() {
		sn = append(sn, n)
	}
	wb.SheetNames = sn
	wb.NumSheets = wb.numSheets()

	// create new excel files to save results to
	xlsxTransformed := excelize.NewFile()
	xlsxRatio := excelize.NewFile()
	xlsxThreshold := excelize.NewFile()
	xlsxSorted := excelize.NewFile()

	// iterate over sheets in workbook
	for i := 0; i < wb.NumSheets; i++ {
		// populate dimension field of excelWorkbook for the current sheet
		wb.Dims = wb.dims(wb.SheetNames[i])

		// print name of current sheet
		fmt.Printf("opened sheet: %s (%d of %d)\n", wb.SheetNames[i], i+1, wb.NumSheets)

		// create a sheet in new workbook with same name to save transformed data
		fmt.Println("creating new sheet to write data to...")
		_ = xlsxTransformed.NewSheet(wb.SheetNames[i])
		_ = xlsxRatio.NewSheet(wb.SheetNames[i])
		_ = xlsxThreshold.NewSheet(wb.SheetNames[i])
		_ = xlsxSorted.NewSheet(wb.SheetNames[i])

		// find the starting index of the actual data matrix
		id, err := wb.startRow(wb.SheetNames[i], "Time (sec)")
		if err != nil {
			fmt.Printf("error while trying to find data: %s\n", err)
			fmt.Println("attempting to analyze data anyways...")
		} else {
			fmt.Printf("found ID: %d --> will start here\n", id)
		}

		// get data
		m := wb.XLSX.GetRows(wb.SheetNames[i])

		// initialize a column counter and a ratio counter
		colCounter := 1
		ratioCounter := 1

		// start analysis
		for j := 1; j < (wb.Dims[1] - 2); j++ { // don't want the last two background columns

			// set column counter and ratio counter to 1 whenever a new worksheet is processed
			if j == 1 {
				colCounter = 1
				ratioCounter = 1
			}

			if mod := j % SKIP; mod == 0 {
				if *verbose {
					fmt.Printf("skipping unwanted column: %d\n", j)
				}
				continue
			}

			// create a column header with the same value as in the original sheet
			currentCol := fmt.Sprintf("%s1", getColumn(colCounter))
			xlsxTransformed.SetCellValue(wb.SheetNames[i], currentCol, m[id][j])

			// verbose output option lets the user see whenever a new column header is written
			if *verbose {
				fmt.Printf("wrote new column header: %v in %s\n", m[id][j], currentCol)
			}

			for k := (id + 1); k < wb.Dims[0]; k++ {

				// offset indicates which background column should be used
				var offset int
				switch {
				case ((j + 1) % 3) == 0:
					offset = 1
				case ((j + 2) % 3) == 0:
					offset = 2 // because go is 0 indexed
				default:
					log.Fatal("something went wrong while performing background corrections")
				}

				// perform background correction of values
				v1, err := strconv.ParseFloat(m[k][j], 64)
				if err != nil {
					log.Fatalf("fatal error converting indices: %s\n", err)
				}
				v2, err := strconv.ParseFloat(m[k][(wb.Dims[1]-offset)], 64)
				if err != nil {
					log.Fatalf("fatal error converting indices: %s\n", err)
				}

				// write corrected value to cell in new workbook (while always starting at row 2, because row 1 holds the labels)
				currentCell := fmt.Sprintf("%s%d", getColumn(colCounter), ((k - id) + 1))
				xlsxTransformed.SetCellValue(wb.SheetNames[i], currentCell, v1-v2)

				// with verbose output, every original and new value will be printed to Stdout
				if *verbose {
					fmt.Printf("default - old value: %v, bg: %v, corrected: %v\n", v1, v2, v1-v2)
				}
			}

			// create a column header for ratios every other column
			if (j % 2) == 0 {

				// write column headers
				currentCol := fmt.Sprintf("%s1", getColumn(ratioCounter))
				currentCell := fmt.Sprintf("cell %d", ratioCounter)
				xlsxRatio.SetCellValue(wb.SheetNames[i], currentCol, currentCell)

				// increment the ratio Counter
				ratioCounter++
			}

			// increment column counter and print current column ONLY if no column is skipped (and verbose output is true)
			if *verbose {
				fmt.Printf("current column: %d\n", colCounter)
			}
			colCounter++
		}

		// done with analysis of one sheet in workbook print summary statistics
		fmt.Printf("summary:\n\tnumber of processed [rows columns]- %v\n\n", wb.Dims)

		// iterate over data in current sheet to create ratios that can be written to xlsxRatio
		// get transformed data
		tm := xlsxTransformed.GetRows(wb.SheetNames[i])

		// continue if current sheet is empty
		if tm == nil || len(tm) < 2 || len(tm[0]) < 2 {
			continue
		}

		// initialize another counter
		rc := 1

		for c := 0; c < len(tm[0]); c += 2 { // iterate over every second column
			for r := 1; r < len(tm); r++ { // iterate over rows starting at row two (row one is header)
				// if r > trimOutput, stop calculating ratios
				if r > *trimOutput {
					if *verbose {
						fmt.Printf("trimmed after %d measurements\n", *trimOutput)
					}
					break
				}
				// string to float conversion
				r1, err := strconv.ParseFloat(tm[r][c], 64)
				if err != nil {
					log.Fatalf("fatal error converting indices: %s\n", err)
				}
				r2, err := strconv.ParseFloat(tm[r][c+1], 64)
				if err != nil {
					log.Fatalf("fatal error converting indices: %s\n", err)
				}

				// get current cell and write
				cl := fmt.Sprintf("%s%d", getColumn(rc), (r + 1)) // need 1 for subsetting but A2 for Excel
				xlsxRatio.SetCellValue(wb.SheetNames[i], cl, (r1 / r2))
				if *verbose {
					fmt.Printf("wrote ratio: %v\n", (r1 / r2))
				}

			}
			rc++
		}

		// add two chart to every ratio data sheet
		// the only purpose of 'shnm' is to reduce the length of the following assignments; don't use it anywhere else
		shnm := wb.SheetNames[i]
		chartSettings1 := fmt.Sprintf("{\"type\":\"line\",\"dimension\":{\"width\":1040,\"height\":640},\"series\":[{\"name\":\"%v!$A$1\",\"values\":\"%v!$A$2:$A$470\"},{\"name\":\"%v!$B$1\",\"values\":\"%v!$B$2:$B$470\"},{\"name\":\"%v!$C$1\",\"values\":\"%v!$C$2:$C$470\"},{\"name\":\"%v!$D$1\",\"values\":\"%v!$D$2:$D$470\"},{\"name\":\"%v!$E$1\",\"values\":\"%v!$E$2:$E$470\"},{\"name\":\"%v!$F$1\",\"values\":\"%v!$F$2:$F$470\"}],\"title\":{\"name\":\"Response Profile\"}}", shnm, shnm, shnm, shnm, shnm, shnm, shnm, shnm, shnm, shnm, shnm, shnm)
		chartSettings2 := fmt.Sprintf("{\"type\":\"line\",\"dimension\":{\"width\":1040,\"height\":640},\"series\":[{\"name\":\"%v!$G$1\",\"values\":\"%v!$G$2:$G$470\"},{\"name\":\"%v!$H$1\",\"values\":\"%v!$H$2:$H$470\"},{\"name\":\"%v!$I$1\",\"values\":\"%v!$I$2:$I$470\"},{\"name\":\"%v!$J$1\",\"values\":\"%v!$J$2:$J$470\"},{\"name\":\"%v!$K$1\",\"values\":\"%v!$K$2:$K$470\"},{\"name\":\"%v!$L$1\",\"values\":\"%v!$L$2:$L$470\"}],\"title\":{\"name\":\"Response Profile\"}}", shnm, shnm, shnm, shnm, shnm, shnm, shnm, shnm, shnm, shnm, shnm, shnm)
		if *addChart {
			xlsxRatio.AddChart(wb.SheetNames[i], "A470", chartSettings1)
			xlsxRatio.AddChart(wb.SheetNames[i], "R470", chartSettings2)
			if *verbose {
				fmt.Printf("added chart to sheet %v with settings: %s\n", wb.SheetNames[i], chartSettings1)
				fmt.Printf("added chart to sheet %v with settings: %s\n", wb.SheetNames[i], chartSettings2)
			}
		}

		// look for peaks with the range of --start (sortStart) and --stop (sortEnd) and sort the ratio columns accordingly
		// use a map to remember the columns that were already copied to the new workbook (xlsxSorted)
		ratioStrings := xlsxRatio.GetRows(wb.SheetNames[i])
		peaks := make(map[int]float64)
		ratioToSort := make([][]float64, 0)

		// parse ratioToSort values into an new slice after converting strings to float64s
		for c := 0; c < len(ratioStrings[0]); c++ {
			// create new slice and append it to a slice of slices
			newArr := make([]float64, len(ratioStrings))

			// initialize an independent value counter
			vc := 0

			// check validity of stop value for search
			var stop int
			if *sortEnd <= len(ratioStrings) {
				stop = *sortEnd
			} else {
				stop = len(ratioStrings)
			}

			// iterate over rows and add all values that are within the sorting range to the slice
			for r := *sortStart; r < stop; r++ {
				val, err := strconv.ParseFloat(ratioStrings[r][c], 64)
				if err != nil {
					log.Fatalf("error while converting indices: %s\n", err)
				}
				if *verbose {
					fmt.Printf("writing %v at [%d][%d]\n", val, r, c)
				}
				newArr[vc] = val
				vc++
			}
			// append new values to slice
			ratioToSort = append(ratioToSort, newArr)
		}

		// iterate over columns of ratioToSort and save to last value of the ordered slice to a map
		for i := 0; i < len(ratioToSort); i++ {
			if *verbose {
				fmt.Printf("sorting column %d\n", i)
			}
			sort.Float64s(ratioToSort[i])
			peaks[i] = ratioToSort[i][len(ratioToSort[0])-1]
		}
		if *verbose {
			fmt.Printf("%+v\n", peaks)
		}

		// print ordered values to screen if flag is set to true; make sure to copy peaks, though!
		tmpMap := make(map[int]float64)
		for key, val := range peaks {
			tmpMap[key] = val
		}
		if *printMap {
			fmt.Printf("ordered values for %s: ", wb.SheetNames[i])
			for {
				if len(tmpMap) == 0 {
					break
				}
				key := findMaxElem(tmpMap)
				fmt.Printf("cell %d: %v ", key+1, tmpMap[key])
				delete(tmpMap, key)
			}
			fmt.Println()
		}

		// return key of max value ==> get that column from ratioToSort ==> write to output ==> delete index from map
		for ii := 0; ii < len(ratioToSort); ii++ {
			// verbose output prints every max map key
			if *verbose {
				fmt.Printf("dim1: %d, dim2: %d\n", len(ratioToSort), len(ratioToSort[0]))
				fmt.Printf("key of current max value in this map: %v\n", findMaxElem(peaks))
			}

			key := findMaxElem(peaks)
			for j := 0; j < len(ratioToSort[0]); j++ {
				// get current cell and write value
				cl := fmt.Sprintf("%s%d", getColumn(ii+1), (j + 1)) // need 0 for subsetting but A2 for Excel
				// write header and continue for j == 0
				if j == 0 {
					xlsxSorted.SetCellValue(wb.SheetNames[i], cl, ratioStrings[j][key])
					continue
				}
				if *verbose {
					fmt.Printf("writing sorted value %v at [%d][%d]\n", ratioStrings[j][key], key, j)
				}
				v, err := strconv.ParseFloat(ratioStrings[j][key], 64)
				if err != nil {
					log.Fatalf("error while converting string: %s\n", err)
				}
				xlsxSorted.SetCellValue(wb.SheetNames[i], cl, v)
			}
			delete(peaks, key)
		}

		// drop columns if not at least one value is > --threshold (this behavior is overriden by --threshold 0)
		if *responseThreshold != 0 {
			// TODO: implement threshold functionality
			// TODO: save thresholded data to a separate file
		}
	}
	printDelim()

	// print some more statistics
	fmt.Printf("summary:\n\tnumber of precessed sheets - %d\n", wb.NumSheets)
	fmt.Printf("\tcreated charts - %v\n", *addChart)
	fmt.Printf("\tsorted ratios in range [lo][hi] - [%d][%d]\n", *sortStart, *sortEnd)
	fmt.Printf("\tratios trimmed after %d measurements\n", *trimOutput)
	if *responseThreshold != 0 {
		fmt.Printf("\tused response threshold: %v\n", *responseThreshold)
	}

	// get current time to create a unique file name
	t := time.Now()
	year, month, day := t.Date()
	hour, min, sec := t.Clock()
	transformedFileName := fmt.Sprintf("%v%v%v_%vh%vmin%vs_transformed_data.xlsx", year, month, day, hour, min, sec)
	ratioFileName := fmt.Sprintf("%v%v%v_%vh%vmin%vs_ratios.xlsx", year, month, day, hour, min, sec)
	sortedRatioFileName := fmt.Sprintf("%v%v%v_%vh%vmin%vs_sorted_ratios.xlsx", year, month, day, hour, min, sec)

	// save output file
	fmt.Printf("writing transformed data to file: %s\n", transformedFileName)
	xlsxTransformed.SaveAs(transformedFileName)
	fmt.Printf("writing ratios to file: %s\n", ratioFileName)
	xlsxRatio.SaveAs(ratioFileName)
	fmt.Printf("writing sorted ratios to file: %s\n", sortedRatioFileName)
	xlsxSorted.SaveAs(sortedRatioFileName)

	// save threshold file
	if *responseThreshold != 0 {
		thresholdFileName := fmt.Sprintf("%v%v%v_%vh%vmin%vs_data_with_threshold.xlsx", year, month, day, hour, min, sec)
		fmt.Printf("writing threshold data to file: %s\n", thresholdFileName)
		xlsxThreshold.SaveAs(thresholdFileName)
	}

}

// prints a useful delimiter
func printDelim() {
	for i := 0; i < 70; i++ {
		fmt.Printf("%s", "-")
		time.Sleep(10 * time.Millisecond)
	}
	fmt.Println()
}

// takes an integer and returns an Excel-style string representation of it (e.g. 1 = A, 3 = C, 27 = AA, ...)
// the current implementation only works for a limited amount of cells, though
func getColumn(num int) string {
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

// helper function for iterating over 'peaks' map;
// find max value ==> get index ==> return index of max value
func findMaxElem(input map[int]float64) int {
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
