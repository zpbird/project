package main

import (
	"fmt"
	"monitor/api"
	// "github.com/360EntSecGroup-Skylar/excelize/v2"
)

func main() {

	// 检查程序文件位置及模板主目录是否存在
	templateDir := api.CheckAppDir()

	// 选择公司
	selCompany := api.GetCompany(templateDir)
	fmt.Printf("%s\n", selCompany)

	// 选择时间段
	selYear, selSmon, selEmon := api.GetTime()
	fmt.Println(selYear, selSmon, selEmon)
}
