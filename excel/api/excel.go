// Package api import "monitor/api" or "excel/api"
package api

import (
	"fmt"
	"os"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/zpbird/zp-go-mod/ztimes"
)

var sysSep = string(os.PathSeparator)

// SumExcel 汇总表结构体
type SumExcel struct {
	SheetList []struct {
		SheetName       string
		Header          string
		Title           map[string]string
		ContentVariable map[string]string
		ContentFixed    map[string]string
		Footer          map[string]string
	}
}

// NewSumExcel ...
func NewSumExcel() *SumExcel {
	return &SumExcel{
		SheetList: make([]struct {
			SheetName       string
			Header          string
			Title           map[string]string
			ContentVariable map[string]string
			ContentFixed    map[string]string
			Footer          map[string]string
		}, 3), // 是否应该默认为0，或指定个数
	}
}

// SxSumExceltmp 世鑫汇总文件模板...
func SxSumExceltmp(dataDir, videoDataDir string, year, mon int) (sx *SumExcel) {
	sx = NewSumExcel()
	// Sheet[0]：汇总
	sx.SheetList[0].SheetName = "汇总"
	sx.SheetList[0].Header = "世鑫录像与地泵数据比对" + strconv.Itoa(year) + "年" + fmt.Sprintf("%02d", mon) + "月"
	sx.SheetList[0].Title = map[string]string{
		"日期":   "A1",
		"地泵数据": "B1",
		"录像数据": "C1",
		"比对结果": "D1",
		"地泵明细": "E1",
		"录像明细": "F1",
	}
	sx.SheetList[0].ContentFixed = map[string]string{
		"日期":   sx.SheetList[0].ContentVariable["contentVar"],
		"地泵数据": "='" + dataDir + sysSep + strconv.Itoa(year) + sysSep + fmt.Sprintf("%02d", mon) + sysSep + "[" + sx.SheetList[0].ContentVariable["contentVar"] + ".xlsx]称重记录'!$C$3",
		"录像数据": "='" + videoDataDir + sysSep + strconv.Itoa(year) + sysSep + fmt.Sprintf("%02d", mon) + sysSep + "[" + sx.SheetList[0].ContentVariable["contentVar"] + ".xlsx]data'!$E$11",
		"比对结果": "",
		"地泵明细": "=HYPERLINK(\"" + dataDir + sysSep + strconv.Itoa(year) + sysSep + fmt.Sprintf("%02d", mon) + sysSep + sx.SheetList[0].ContentVariable["contentVar"] + ".xlsx\",\"" + sx.SheetList[0].ContentVariable["contentVar"] + "\")",
		"录像明细": "=HYPERLINK(\"" + videoDataDir + sysSep + strconv.Itoa(year) + sysSep + fmt.Sprintf("%02d", mon) + sysSep + "[" + sx.SheetList[0].ContentVariable["contentVar"] + ".xlsx\",\"" + sx.SheetList[0].ContentVariable["contentVar"] + "\")",
	}
	sx.SheetList[0].ContentVariable = map[string]string{
		"dayIndex":   "", // 应该必须最先赋值
		"contentVar": strconv.Itoa(year) + "-" + fmt.Sprintf("%02d", mon) + "-" + sx.SheetList[0].ContentVariable["dayIndex"],
	}
	sx.SheetList[0].Footer = map[string]string{
		"日期":   "汇总",
		"地泵数据": "=\"总车数：\"&tmp!B34&\"  水渣：\"&tmp!C34&\"  矿粉：\"&tmp!D34&\"  其他：\"&tmp!E34",
		"录像数据": "=\"总车数：\"&tmp!G34&\"  水渣：\"&tmp!H34&\"  矿粉：\"&tmp!I34&\"  其他：\"&tmp!J34&\"  异常：\"&tmp!K34",
	}

	// Sheet[1]：tmp

	// Sheet[2]：地泵汇总

	return
}

// MakeSumExcel ...
// selCompany, dataDir, videoDataDir string, year, mon int
func MakeSumExcel() (sumExcelFile *excelize.File, err error) {
	ztimes.GetMonDays(2020, 2)
	sxTmp := SxSumExceltmp("aaa", "bbb", 2019, 5)
	sumExcelFile = excelize.NewFile()
	// 设置页眉
	err = sumExcelFile.SetHeaderFooter("Sheet1", &excelize.FormatHeaderFooter{
		DifferentFirst: true,
		FirstHeader:    `&C` + `&B` + `&16` + `&"微软雅黑,常规"` + sxTmp.SheetList[0].Header,
	})
	if err != nil {
		return
	}
	err = sumExcelFile.SaveAs("./files/ttt.xlsx")
	return

}
