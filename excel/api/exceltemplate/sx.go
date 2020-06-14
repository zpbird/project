package exceltemplate

import (
	"fmt"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

// SxSumExcelTmp 世鑫汇总文件模板...
func SxSumExcelTmp(year, mon int) (sx *SumExcelTmp) {
	sx = NewSumExcelTmp()
	// Sheet[0]：汇总
	sx.SheetList[0].SheetName = "汇总"
	sx.SheetList[0].Header = "世鑫录像与地泵数据比对" + strconv.Itoa(year) + "年" + fmt.Sprintf("%02d", mon) + "月"
	sx.SheetList[0].Title = map[string][2]string{
		"日期":   {"A", "1"},
		"地泵数据": {"B", "1"},
		"录像数据": {"C", "1"},
		"比对结果": {"D", "1"},
		"地泵明细": {"E", "1"},
		"录像明细": {"F", "1"},
	}
	sx.SheetList[0].ContentFixed = map[string]string{
		"日期":   sx.SheetList[0].ContentVariable["contentVar"],
		"地泵数据": "='.." + sysSep + "data" + sysSep + strconv.Itoa(year) + sysSep + fmt.Sprintf("%02d", mon) + sysSep + "[" + sx.SheetList[0].ContentVariable["contentVar"] + ".xlsx]称重记录'!$C$3",
		"录像数据": "='.." + sysSep + "video_data" + sysSep + strconv.Itoa(year) + sysSep + fmt.Sprintf("%02d", mon) + sysSep + "[" + sx.SheetList[0].ContentVariable["contentVar"] + ".xlsx]data'!$E$11",
		"比对结果": "",
		"地泵明细": "=HYPERLINK(\"" + ".." + sysSep + "data" + sysSep + strconv.Itoa(year) + sysSep + fmt.Sprintf("%02d", mon) + sysSep + sx.SheetList[0].ContentVariable["contentVar"] + ".xlsx\",\"" + sx.SheetList[0].ContentVariable["contentVar"] + "\")",
		"录像明细": "=HYPERLINK(\"" + ".." + sysSep + "video_data" + sysSep + strconv.Itoa(year) + sysSep + fmt.Sprintf("%02d", mon) + sysSep + "[" + sx.SheetList[0].ContentVariable["contentVar"] + ".xlsx\",\"" + sx.SheetList[0].ContentVariable["contentVar"] + "\")",
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

// SxMakeSumExcelFile ...
func SxMakeSumExcelFile(year, mon int) (sxSumExcelFile *excelize.File, err error) {
	var style int
	// monDays := ztimes.GetMonDays(year, mon)
	sxTmp := SxSumExcelTmp(year, mon)
	sxSumExcelFile = excelize.NewFile()

	// 设置工作簿默认字体
	sxSumExcelFile.SetDefaultFont(styleDefaultFont)

	// 新建"汇总Sheet"
	indexHz := sxSumExcelFile.NewSheet("汇总")
	sxSumExcelFile.SetActiveSheet(indexHz)
	sxSumExcelFile.DeleteSheet("Sheet1") // 删除默认Sheet1

	// 设置"汇总Sheet"页眉
	err = sxSumExcelFile.SetHeaderFooter("汇总", &excelize.FormatHeaderFooter{
		DifferentFirst: true,
		FirstHeader:    `&C` + `&B` + `&16` + `&"微软雅黑,常规"` + sxTmp.SheetList[0].Header,
	})
	if err != nil {
		fmt.Println(err)
		return
	}

	// 设置"汇总Sheet"标题行
	for key, value := range sxTmp.SheetList[0].Title {
		if err = sxSumExcelFile.SetCellValue(sxTmp.SheetList[0].SheetName, value[0]+value[1], key); err != nil {
			fmt.Println(err)
			return
		} else {
			style, err = sxSumExcelFile.NewStyle(styleTitle)
			if err != nil {
				fmt.Println(err)
				return
			}
			sxSumExcelFile.SetCellStyle(sxTmp.SheetList[0].SheetName, value[0]+value[1], value[0]+value[1], style)

		}

		switch key {
		case "日期":
			if err = sxSumExcelFile.SetColWidth(sxTmp.SheetList[0].SheetName, value[0], value[0], 16); err != nil {
				fmt.Println(err)
				return
			}
		case "地泵数据":
			if err = sxSumExcelFile.SetColWidth(sxTmp.SheetList[0].SheetName, value[0], value[0], 40); err != nil {
				fmt.Println(err)
				return
			}
		case "录像数据":
			if err = sxSumExcelFile.SetColWidth(sxTmp.SheetList[0].SheetName, value[0], value[0], 48); err != nil {
				fmt.Println(err)
				return
			}
		case "比对结果":
			if err = sxSumExcelFile.SetColWidth(sxTmp.SheetList[0].SheetName, value[0], value[0], 36); err != nil {
				fmt.Println(err)
				return
			}
		case "地泵明细":
			if err = sxSumExcelFile.SetColWidth(sxTmp.SheetList[0].SheetName, value[0], value[0], 16); err != nil {
				fmt.Println(err)
				return
			}
		case "录像明细":
			if err = sxSumExcelFile.SetColWidth(sxTmp.SheetList[0].SheetName, value[0], value[0], 16); err != nil {
				fmt.Println(err)
				return
			}
		}

	}

	// 保存文件
	err = sxSumExcelFile.SaveAs("./files/ttt.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	return

}
