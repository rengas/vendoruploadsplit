package main

import (
	"flag"
	"fmt"

	"github.com/xuri/excelize/v2"
)

var (
	filename     = flag.String("filename", "vendors.xlsx", "file to split")
	rowsPersheet = flag.Int("rows", 10000, "number of rows per sheet in new file")
)

func main() {
	flag.Parse()
	
	f, err := excelize.OpenFile(*filename)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	rows, err := f.GetRows(f.GetSheetName(0))
	if err != nil {
		fmt.Println(err)
		return
	}

	if len(rows) > 0 {
		rows = rows[1:]
		var file *excelize.File

		file = excelize.NewFile()

		err = file.SetCellValue("sheet1", fmt.Sprintf("A%d", 1), "vendor_code")
		if err != nil {
			fmt.Println(err)
			return
		}
		err = file.SetCellValue("sheet1", fmt.Sprintf("B%d", 1), "add")
		if err != nil {
			fmt.Println(err)
			return
		}
		countFile := 0
		rowReset := 1
		var chunks [][][]string
		chunkSize := *rowsPersheet
		for i := 0; i < len(rows); i += chunkSize {
			end := i + chunkSize
			// necessary check to avoid slicing beyond
			// slice capacity
			if end > len(rows) {
				end = len(rows)
			}
			chunks = append(chunks, rows[i:end])
		}

		for i := 0; i < len(chunks); i++ {
			rows := chunks[i]
			file = excelize.NewFile()
			rowReset = 1
			err = file.SetCellValue("sheet1", fmt.Sprintf("A%d", 1), "vendor_code")
			if err != nil {
				panic(err)
			}
			err = file.SetCellValue("sheet1", fmt.Sprintf("B%d", 1), "add")
			if err != nil {
				panic(err)
			}

			for _, row := range rows {
				rowReset += 1
				for i, colCell := range row {
					if i == 0 {
						err := file.SetCellValue("sheet1", fmt.Sprintf("A%d", rowReset), colCell)
						if err != nil {
							fmt.Println(err)
							return
						}
					}
					if i == 1 {
						err := file.SetCellValue("sheet1", fmt.Sprintf("B%d", rowReset), colCell)
						if err != nil {
							fmt.Println(err)
							return
						}
					}
				}
			}

			err := file.SaveAs(fmt.Sprintf("vendorCode_%d.xlsx", countFile))
			if err != nil {
				fmt.Println(err)
				return
			}

			err = file.Close()
			if err != nil {
				fmt.Println(err)
				return
			}
			countFile += 1
		}
	}
}
