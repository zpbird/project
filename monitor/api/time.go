package api

import (
	"fmt"
	"strconv"
)

// GetTime 返回用户选择的时间...
func GetTime() (year int, sMon int, eMon int) {
	// 初始值
	year = 0
	sMon = 0
	eMon = 0

	// 获取年份
	for year == 0 {
		Clear() // 清屏
		fmt.Printf("\n\n")
		fmt.Print("请输入年份[2018 - 9999] ：")

		// 验证输入的字符串为数字形式(一定程度上)
		var getStrYear string
		fmt.Scanln(&getStrYear)
		if i, err := strconv.ParseInt(getStrYear, 0, 64); err != nil {
			fmt.Println("输入时发生错误！")
			year = 0
		} else if year = int(i); year < 2018 || year > 9999 {
			year = 0
		} else {
			year = int(i)
		}
	}

	// 获取起止月份
	for sMon == 0 || eMon == 0 {
		Clear() // 清屏
		fmt.Printf("\n\n")
		fmt.Print("按顺序输入起始和结束月份，用空格分开[1  12] ：")
		// 验证输入的字符串为数字形式(一定程度上)
		var getStrSmon, getStrEmon string
		fmt.Scanln(&getStrSmon, &getStrEmon)

		// 验证开始月份
		if i, err := strconv.ParseInt(getStrSmon, 0, 64); err != nil {
			fmt.Println("输入时发生错误！")
			sMon = 0
		} else if sMon = int(i); sMon < 1 || sMon > 12 {
			sMon = 0
		} else {
			sMon = int(i)
		}

		// 验证结束月份
		if i, err := strconv.ParseInt(getStrEmon, 0, 64); err != nil {
			fmt.Println("输入时发生错误！")
			eMon = 0
		} else if eMon = int(i); eMon < 1 || eMon > 12 {
			eMon = 0
		} else if sMon > eMon {
			eMon = 0
		} else {
			eMon = int(i)
		}

	}

	return year, sMon, eMon
}
