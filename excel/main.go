package main

import (
	"excel/api/exceltemplate"
	"fmt"
)

// main ...
func main() {
	sxexcel, _ := exceltemplate.SxMakeSumExcelFile(2019, 7)

	err := sxexcel.SaveAs("./files/ttt.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

}
