package main

import "github.com/360EntSecGroup-Skylar/excelize"

func main() {
	f := excelize.NewFile()
	// 创建一个工作表
	index := f.NewSheet("Sheet2")
	// 设置单元格的值
	f.SetCellValue("Sheet2", "A2", "Hello world.aaaaaaaaaaaaaaaaaaaa")
	f.SetCellValue("Sheet1", "B2", 100)
	// 设置工作簿的默认工作表
	f.SetActiveSheet(index)
	// 根据指定路径保存文件
	if err := f.SaveAs("Book1.xlsx"); err != nil {
		println(err.Error())
		// aaa
	}
	fff := excelize.NewFile()
	fff.NewSheet("aaaaaaaaaaaaaaaaaaaa")
	ccc := excelize.NewFile()
	ccc.NewSheet("asdf")
}

// aaa ...
func aaa() {

}
