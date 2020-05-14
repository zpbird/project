package main

import (
	"fmt"
	"monitor/api"
	"strconv"
	// "github.com/360EntSecGroup-Skylar/excelize/v2"
)

func main() {

	// 检查程序文件位置及模板主目录是否存在
	templateDir := api.CheckAppDir()

reInput: // 标签:重新输入

	// 选择公司
	selCompany := api.GetCompany(templateDir)
	fmt.Printf("%s\n", selCompany)

	// 选择时间段
	selYear, selSmon, selEmon := api.GetTime()
	fmt.Println(selYear, selSmon, selEmon)

	// 确认选择信息
	for acceptSel := true; acceptSel; {
		monSection := ""
		if selEmon-selSmon == 0 {
			monSection = strconv.Itoa(selEmon)
		} else {
			monSection = strconv.Itoa(selSmon) + "-" + strconv.Itoa(selEmon)
		}

		api.Clear()
		fmt.Printf("\n\n")
		fmt.Printf("公司：%s     ", selCompany)
		fmt.Printf("年月份：%d年%s月份\n\n", selYear, monSection)
		fmt.Printf("1：确认  ")
		fmt.Printf("2：重新选择\n\n")
		fmt.Printf("请输入选择[1 或 2]：")

		var getStr string
		fmt.Scanln(&getStr)
		if i, err := strconv.ParseInt(getStr, 0, 64); err != nil {
			fmt.Println("输入时发生错误！")
		} else if int(i) == 1 {
			acceptSel = false
		} else if int(i) == 2 {
			goto reInput
		}
	}

	fmt.Println(selCompany, selYear, selSmon, selEmon)
}
