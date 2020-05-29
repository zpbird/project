package main

import (
	"fmt"
	"monitor/api"
	"strconv"

	"github.com/zpbird/zp-go-mod/zinput"
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

		zinput.Clear()

		tTermStr := "公司：" + selCompany + "     " + "年月份：" + strconv.Itoa(selYear) + "年" + monSection + "月份\n\n" + "y：确认  " + "n：重新选择\n\n" + "请输入选择[y 或 n]："
		n := zinput.Input(tTermStr, zinput.RegYn)
		if n == "y" {
			acceptSel = false
		} else if n == "n" {
			goto reInput
		}

	}

	fmt.Println(selCompany, selYear, selSmon, selEmon)
}
