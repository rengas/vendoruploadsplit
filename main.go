package main

import (
	"fmt"
	"math"

	"github.com/xuri/excelize/v2"
)

func main() {
	f, err := excelize.OpenFile("vendors.xlsx")
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

	rows, err := f.GetRows("result")
	if err != nil {
		fmt.Println(err)
		return
	}

	if len(rows) > 0 {

		rows = rows[1:]

		totalSheets := int(math.Ceil(float64(len(rows)/10000))) + 2

		fmt.Println(totalSheets)

		var file *excelize.File

		file = excelize.NewFile()

		err = file.SetCellValue("sheet1", fmt.Sprintf("A%d", 1), "vendor_code")
		if err != nil {
			panic(err)
		}
		err = file.SetCellValue("sheet1", fmt.Sprintf("B%d", 1), "add")
		if err != nil {
			panic(err)
		}
		countFile := 0
		rowReset := 1
		var chunks [][][]string
		chunkSize := 10000
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
							panic(err)
						}
					}
					if i == 1 {
						err := file.SetCellValue("sheet1", fmt.Sprintf("B%d", rowReset), colCell)
						if err != nil {
							panic(err)
						}
					}
				}
			}

			err := file.SaveAs(fmt.Sprintf("vendorCode_%d.xlsx", countFile))
			if err != nil {
				panic(err)
			}

			err = file.Close()
			if err != nil {
				panic(err)
			}
			countFile += 1
		}
	}
}
