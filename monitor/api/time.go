package api

import (
	"fmt"
	"strconv"

	"github.com/zpbird/zp-go-mod/zinput"
)

// GetTime 返回用户选择的时间...
func GetTime() (year int, sMon int, eMon int) {
	// 初始值
	year = 0
	sMon = 0
	eMon = 0

	// 获取年份
	zinput.Clear()
	year, _ = strconv.Atoi(zinput.Input("\"年份\"，"+zinput.RegYearRul, zinput.RegYear))

	// 获取起止月份
	trigger := true
	for trigger {
		zinput.Clear()
		sMon, _ = strconv.Atoi(zinput.Input("\"开始月份\"，"+zinput.RegMonRule, zinput.RegMon))
		zinput.Clear()
		eMon, _ = strconv.Atoi(zinput.Input("\"结束月份\"，"+zinput.RegMonRule, zinput.RegMon))
		trigger = false
		if eMon < sMon {
			zinput.Clear()
			fmt.Printf("\"结束月份\"必须大于或等于\"开始月份\"，目前输入的值为：\"开始月份\"[%d] \"结束月份\"[%d]\n", sMon, eMon)
			zinput.StopContinue()
			trigger = true
		}
	}

	return year, sMon, eMon
}
