package main

import (
	"bytes"
	"embed"
	"flag"
	"fmt"

	"github.com/xuri/excelize/v2"
)

//go:embed  template_add.xlsx template_remove.xlsx
var content embed.FS

var (
	filename      = flag.String("filename", "vendors.xlsx", "file to split")
	rowsPersheet  = flag.Int("rows", 10000, "number of rows per sheet in new file")
	paymentMethod = flag.String("payment-method", "default", "number of rows per sheet in new file")
	operation     = flag.String("operation", "add", "add/remove payment method for vendor code")
)

func main() {
	flag.Parse()

	var templateFile = "template_add.xlsx"
	if *operation == "remove" {
		templateFile = "template_remove.xlsx"
	}

	data, err := content.ReadFile(templateFile)
	if err != nil {
		fmt.Println(err)
		return
	}

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
			file, err = excelize.OpenReader(bytes.NewReader(data))
			if err != nil {
				fmt.Println(err)
				return
			}
			rowReset = 1

			for _, row := range rows {
				rowReset += 1
				for i, colCell := range row {
					if i == 0 {
						err := file.SetCellValue("VendorPaymentTypes", fmt.Sprintf("A%d", rowReset), colCell)
						if err != nil {
							fmt.Println(err)
							return
						}
						err = file.SetCellValue("VendorPaymentTypes", fmt.Sprintf("B%d", rowReset), *paymentMethod)
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
	} else {
		fmt.Printf("%s is empty", *filename)
	}
}
