package exceltemplate

import (
	"fmt"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/zpbird/zp-go-mod/ztimes"
)

// SxSumExcelTmp 世鑫汇总文件模板...
func SxSumExcelTmp(year, mon int) (sx *SumExcelTmp) {
	sx = NewSumExcelTmp()
	// Sheet[0]：汇总
	sx.SheetList[0].SheetName = "汇总"
	sx.SheetList[0].Header = "世鑫录像与地泵数据比对" + strconv.Itoa(year) + "年" + fmt.Sprintf("%02d", mon) + "月"
	sx.SheetList[0].Title = map[string][]string{
		"日期":   {"A", "1"},
		"地泵数据": {"B", "1"},
		"录像数据": {"C", "1"},
		"比对结果": {"D", "1"},
		"地泵明细": {"E", "1"},
		"录像明细": {"F", "1"},
	}
	sx.SheetList[0].ContentVariable = map[string]string{
		"contentVarInit": strconv.Itoa(year) + "-" + fmt.Sprintf("%02d", mon) + "-",
		"contentVar":     "",
	}

	sx.SheetList[0].ContentFixed = map[string]string{
		"日期":   sx.SheetList[0].ContentVariable["contentVar"],
		"地泵数据": "='.." + sysSep + "data" + sysSep + strconv.Itoa(year) + sysSep + fmt.Sprintf("%02d", mon) + sysSep + "[" + sx.SheetList[0].ContentVariable["contentVar"] + ".xls]称重记录'!$C$3",
		"录像数据": "='.." + sysSep + "video_data" + sysSep + strconv.Itoa(year) + sysSep + fmt.Sprintf("%02d", mon) + sysSep + "[" + sx.SheetList[0].ContentVariable["contentVar"] + ".xlsx]data'!$E$11",
		"比对结果": "",
		"地泵明细": "=HYPERLINK(\"" + ".." + sysSep + "data" + sysSep + strconv.Itoa(year) + sysSep + fmt.Sprintf("%02d", mon) + sysSep + sx.SheetList[0].ContentVariable["contentVar"] + ".xls\",\"" + sx.SheetList[0].ContentVariable["contentVar"] + "\")",
		"录像明细": "=HYPERLINK(\"" + ".." + sysSep + "video_data" + sysSep + strconv.Itoa(year) + sysSep + fmt.Sprintf("%02d", mon) + sysSep + sx.SheetList[0].ContentVariable["contentVar"] + ".xlsx\",\"" + sx.SheetList[0].ContentVariable["contentVar"] + "\")",
	}

	sx.SheetList[0].Footer = map[string]string{
		"日期":   "汇总",
		"地泵数据": "=\"总车数：\"&tmp!B34&\"  水渣：\"&tmp!C34&\"  矿粉：\"&tmp!D34&\"  其他：\"&tmp!E34",
		"录像数据": "=\"总车数：\"&tmp!G34&\"  水渣：\"&tmp!H34&\"  矿粉：\"&tmp!I34&\"  其他：\"&tmp!J34&\"  异常：\"&tmp!K34",
	}

	// Sheet[1]：tmp
	sx.SheetList[1].SheetName = "tmp"
	sx.SheetList[1].Header = "世鑫录像与地泵数据比对" + strconv.Itoa(year) + "年" + fmt.Sprintf("%02d", mon) + "月"
	sx.SheetList[1].Title = map[string][]string{
		"地泵数据":  {"A1", "E1"},
		"录像数据":  {"G1", "K1"},
		"日期":    {"日期", "A", "2"},
		"地泵总车数": {"总车数", "B", "2"},
		"地泵水渣":  {"水渣", "C", "2"},
		"地泵矿粉":  {"矿粉", "D", "2"},
		"地泵其他":  {"其他", "E", "2"},
		"录像总车数": {"总车数", "G", "2"},
		"录像水渣":  {"水渣", "H", "2"},
		"录像矿粉":  {"矿粉", "I", "2"},
		"录像其他":  {"其他", "J", "2"},
		"录像异常":  {"异常", "K", "2"},
	}
	sx.SheetList[1].ContentVariable = map[string]string{
		"dayInit": strconv.Itoa(year) + "-" + fmt.Sprintf("%02d", mon) + "-",
		"day":     "",
	}
	sx.SheetList[1].ContentFixed = map[string]string{
		"日期":    "",
		"地泵总车数": "=VALUE(MID(汇总!$B" + "varAxis" + ",FIND(\"总车数：\",汇总!$B" + "varAxis" + ",1)+4,FIND(\" \",汇总!$B" + "varAxis" + ",FIND(\"总车数：\",汇总!$B" + "varAxis" + ",1)+4)-(FIND(\"总车数：\",汇总!$B" + "varAxis" + ",1)+4)))",
		"地泵水渣":  "=VALUE(MID(汇总!$B" + "varAxis" + ",FIND(\"水渣：\",汇总!$B" + "varAxis" + ",1)+3,FIND(\" \",汇总!$B" + "varAxis" + ",FIND(\"水渣：\",汇总!$B" + "varAxis" + ",1)+3)-(FIND(\"水渣：\",汇总!$B" + "varAxis" + ",1)+3)))",
		"地泵矿粉":  "=VALUE(MID(汇总!$B" + "varAxis" + ",FIND(\"矿粉：\",汇总!$B" + "varAxis" + ",1)+3,FIND(\" \",汇总!$B" + "varAxis" + ",FIND(\"矿粉：\",汇总!$B" + "varAxis" + ",1)+3)-(FIND(\"矿粉：\",汇总!$B" + "varAxis" + ",1)+3)))",
		"地泵其他":  "=VALUE(MID(汇总!$B" + "varAxis" + ",FIND(\"其他：\",汇总!$B" + "varAxis" + ",1)+3,3))",
		"录像总车数": "=VALUE(MID(汇总!$C" + "varAxis" + ",FIND(\"总车数：\",汇总!$C" + "varAxis" + ",1)+4,FIND(\" \",汇总!$C" + "varAxis" + ",FIND(\"总车数：\",汇总!$C" + "varAxis" + ",1)+4)-(FIND(\"总车数：\",汇总!$C" + "varAxis" + ",1)+4)))",
		"录像水渣":  "=VALUE(MID(汇总!$C" + "varAxis" + ",FIND(\"水渣：\",汇总!$C" + "varAxis" + ",1)+3,FIND(\" \",汇总!$C" + "varAxis" + ",FIND(\"水渣：\",汇总!$C" + "varAxis" + ",1)+3)-(FIND(\"水渣：\",汇总!$C" + "varAxis" + ",1)+3)))",
		"录像矿粉":  "=VALUE(MID(汇总!$C" + "varAxis" + ",FIND(\"矿粉：\",汇总!$C" + "varAxis" + ",1)+3,FIND(\" \",汇总!$C" + "varAxis" + ",FIND(\"矿粉：\",汇总!$C" + "varAxis" + ",1)+3)-(FIND(\"矿粉：\",汇总!$C" + "varAxis" + ",1)+3)))",
		"录像其他":  "=VALUE(MID(汇总!$C" + "varAxis" + ",FIND(\"其他：\",汇总!$C" + "varAxis" + ",1)+3,FIND(\" \",汇总!$C" + "varAxis" + ",FIND(\"其他：\",汇总!$C" + "varAxis" + ",1)+3)-(FIND(\"其他：\",汇总!$C" + "varAxis" + ",1)+3)))",
		"录像异常":  "=VALUE(MID(汇总!$C" + "varAxis" + ",FIND(\"异常：\",汇总!$C" + "varAxis" + ",1)+3,3))",
	}
	sx.SheetList[1].Footer = map[string]string{
		"日期":      "汇总",
		"地泵总车数汇总": "=SUM(B3:B" + "varAxis" + ")",
		"地泵水渣汇总":  "=SUM(C3:C" + "varAxis" + ")",
		"地泵矿粉汇总":  "=SUM(D3:D" + "varAxis" + ")",
		"地泵其他汇总":  "=SUM(E3:E" + "varAxis" + ")",
		"录像总车数汇总": "=SUM(G3:G" + "varAxis" + ")",
		"录像水渣汇总":  "=SUM(H3:H" + "varAxis" + ")",
		"录像矿粉汇总":  "=SUM(I3:I" + "varAxis" + ")",
		"录像其他汇总":  "=SUM(J3:J" + "varAxis" + ")",
		"录像异常汇总":  "=SUM(K3:K" + "varAxis" + ")",
	}
	// Sheet[2]：地泵汇总

	return
}

// SxMakeSumExcelFile ...
func SxMakeSumExcelFile(year, mon int) (sxSumExcelFile *excelize.File, err error) {
	var style int
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

	// 设置"汇总Sheet"内容行
	monDays := ztimes.GetMonDays(year, mon)
	for i := 1; i <= monDays; i++ {
		sxTmp.SheetList[0].ContentVariable["contentVar"] = sxTmp.SheetList[0].ContentVariable["contentVarInit"] + fmt.Sprintf("%02d", i)

		// 日期列
		sxTmp.SheetList[0].ContentFixed["日期"] = sxTmp.SheetList[0].ContentVariable["contentVar"]
		if err = sxSumExcelFile.SetCellValue(sxTmp.SheetList[0].SheetName, sxTmp.SheetList[0].Title["日期"][0]+fmt.Sprintf("%d", i+1), sxTmp.SheetList[0].ContentFixed["日期"]); err != nil {
			fmt.Println(err)
			return
		} else {
			style, err = sxSumExcelFile.NewStyle(styleContentAlignCenter)
			if err != nil {
				fmt.Println(err)
				return
			}
			sxSumExcelFile.SetCellStyle(sxTmp.SheetList[0].SheetName, sxTmp.SheetList[0].Title["日期"][0]+fmt.Sprintf("%d", i+1), sxTmp.SheetList[0].Title["日期"][0]+fmt.Sprintf("%d", i+1), style)
		}

		// 地泵数据列
		sxTmp.SheetList[0].ContentFixed["地泵数据"] = "='.." + sysSep + "data" + sysSep + strconv.Itoa(year) + sysSep + fmt.Sprintf("%02d", mon) + sysSep + "[" + sxTmp.SheetList[0].ContentVariable["contentVar"] + ".xls]称重记录'!$C$3"
		if err = sxSumExcelFile.SetCellFormula(sxTmp.SheetList[0].SheetName, sxTmp.SheetList[0].Title["地泵数据"][0]+fmt.Sprintf("%d", i+1), sxTmp.SheetList[0].ContentFixed["地泵数据"]); err != nil {
			fmt.Println(err)
			return
		} else {
			style, err = sxSumExcelFile.NewStyle(styleContent)
			if err != nil {
				fmt.Println(err)
				return
			}
			sxSumExcelFile.SetCellStyle(sxTmp.SheetList[0].SheetName, sxTmp.SheetList[0].Title["地泵数据"][0]+fmt.Sprintf("%d", i+1), sxTmp.SheetList[0].Title["地泵数据"][0]+fmt.Sprintf("%d", i+1), style)
		}

		// 录像数据列
		sxTmp.SheetList[0].ContentFixed["录像数据"] = "='.." + sysSep + "video_data" + sysSep + strconv.Itoa(year) + sysSep + fmt.Sprintf("%02d", mon) + sysSep + "[" + sxTmp.SheetList[0].ContentVariable["contentVar"] + ".xlsx]data'!$E$11"
		if err = sxSumExcelFile.SetCellFormula(sxTmp.SheetList[0].SheetName, sxTmp.SheetList[0].Title["录像数据"][0]+fmt.Sprintf("%d", i+1), sxTmp.SheetList[0].ContentFixed["录像数据"]); err != nil {
			fmt.Println(err)
			return
		} else {
			style, err = sxSumExcelFile.NewStyle(styleContent)
			if err != nil {
				fmt.Println(err)
				return
			}
			sxSumExcelFile.SetCellStyle(sxTmp.SheetList[0].SheetName, sxTmp.SheetList[0].Title["录像数据"][0]+fmt.Sprintf("%d", i+1), sxTmp.SheetList[0].Title["录像数据"][0]+fmt.Sprintf("%d", i+1), style)
		}

		// 比对结果列
		if err = sxSumExcelFile.SetCellFormula(sxTmp.SheetList[0].SheetName, sxTmp.SheetList[0].Title["比对结果"][0]+fmt.Sprintf("%d", i+1), sxTmp.SheetList[0].ContentFixed["比对结果"]); err != nil {
			fmt.Println(err)
			return
		} else {
			style, err = sxSumExcelFile.NewStyle(styleContent)
			if err != nil {
				fmt.Println(err)
				return
			}
			sxSumExcelFile.SetCellStyle(sxTmp.SheetList[0].SheetName, sxTmp.SheetList[0].Title["比对结果"][0]+fmt.Sprintf("%d", i+1), sxTmp.SheetList[0].Title["比对结果"][0]+fmt.Sprintf("%d", i+1), style)
		}

		// 地泵明细列
		sxTmp.SheetList[0].ContentFixed["地泵明细"] = "=HYPERLINK(\"" + ".." + sysSep + "data" + sysSep + strconv.Itoa(year) + sysSep + fmt.Sprintf("%02d", mon) + sysSep + sxTmp.SheetList[0].ContentVariable["contentVar"] + ".xls\",\"" + sxTmp.SheetList[0].ContentVariable["contentVar"] + "\")"
		if err = sxSumExcelFile.SetCellFormula(sxTmp.SheetList[0].SheetName, sxTmp.SheetList[0].Title["地泵明细"][0]+fmt.Sprintf("%d", i+1), sxTmp.SheetList[0].ContentFixed["地泵明细"]); err != nil {
			fmt.Println(err)
			return
		} else {
			style, err = sxSumExcelFile.NewStyle(&excelize.Style{
				Font: &excelize.Font{
					Color: "003366", Bold: true},
				Alignment: styleAlignCenter,
				Border:    sytleBorderAll,
			})
			if err != nil {
				fmt.Println(err)
				return
			}
			sxSumExcelFile.SetCellStyle(sxTmp.SheetList[0].SheetName, sxTmp.SheetList[0].Title["地泵明细"][0]+fmt.Sprintf("%d", i+1), sxTmp.SheetList[0].Title["地泵明细"][0]+fmt.Sprintf("%d", i+1), style)
		}

		// 录像明细列
		sxTmp.SheetList[0].ContentFixed["录像明细"] = "=HYPERLINK(\"" + ".." + sysSep + "video_data" + sysSep + strconv.Itoa(year) + sysSep + fmt.Sprintf("%02d", mon) + sysSep + sxTmp.SheetList[0].ContentVariable["contentVar"] + ".xlsx\",\"" + sxTmp.SheetList[0].ContentVariable["contentVar"] + "\")"
		if err = sxSumExcelFile.SetCellFormula(sxTmp.SheetList[0].SheetName, sxTmp.SheetList[0].Title["录像明细"][0]+fmt.Sprintf("%d", i+1), sxTmp.SheetList[0].ContentFixed["录像明细"]); err != nil {
			fmt.Println(err)
			return
		} else {
			style, err = sxSumExcelFile.NewStyle(&excelize.Style{
				Font: &excelize.Font{
					Color: "003366", Bold: true},
				Alignment: styleAlignCenter,
				Border:    sytleBorderAll,
			})
			if err != nil {
				fmt.Println(err)
				return
			}
			sxSumExcelFile.SetCellStyle(sxTmp.SheetList[0].SheetName, sxTmp.SheetList[0].Title["录像明细"][0]+fmt.Sprintf("%d", i+1), sxTmp.SheetList[0].Title["录像明细"][0]+fmt.Sprintf("%d", i+1), style)
		}
	}

	// 设置汇总Sheet 汇总行
	sxTmp.SheetList[0].Footer["地泵数据"] = "=\"总车数：\"&tmp!B" + fmt.Sprintf("%d", monDays+3) + "&\"  水渣：\"&tmp!C" + fmt.Sprintf("%d", monDays+3) + "&\"  矿粉：\"&tmp!D" + fmt.Sprintf("%d", monDays+3) + "&\"  其他：\"&tmp!E" + fmt.Sprintf("%d", monDays+3)
	sxTmp.SheetList[0].Footer["录像数据"] = "=\"总车数：\"&tmp!G" + fmt.Sprintf("%d", monDays+3) + "&\"  水渣：\"&tmp!H" + fmt.Sprintf("%d", monDays+3) + "&\"  矿粉：\"&tmp!I" + fmt.Sprintf("%d", monDays+3) + "&\"  其他：\"&tmp!J" + fmt.Sprintf("%d", monDays+3) + "&\"  异常：\"&tmp!K" + fmt.Sprintf("%d", monDays+3)
	for key, value := range sxTmp.SheetList[0].Footer {
		switch key {
		case "日期":
			if err = sxSumExcelFile.SetCellValue(sxTmp.SheetList[0].SheetName, sxTmp.SheetList[0].Title["日期"][0]+fmt.Sprintf("%d", monDays+2), value); err != nil {
				fmt.Println(err)
				return
			} else {
				style, err = sxSumExcelFile.NewStyle(styleTitle)
				if err != nil {
					fmt.Println(err)
					return
				}
				sxSumExcelFile.SetCellStyle(sxTmp.SheetList[0].SheetName, sxTmp.SheetList[0].Title["日期"][0]+fmt.Sprintf("%d", monDays+2), sxTmp.SheetList[0].Title["日期"][0]+fmt.Sprintf("%d", monDays+2), style)
			}
		case "地泵数据":
			if err = sxSumExcelFile.SetCellFormula(sxTmp.SheetList[0].SheetName, sxTmp.SheetList[0].Title["地泵数据"][0]+fmt.Sprintf("%d", monDays+2), value); err != nil {
				fmt.Println(err)
				return
			} else {
				style, err = sxSumExcelFile.NewStyle(styleContent)
				if err != nil {
					fmt.Println(err)
					return
				}
				sxSumExcelFile.SetCellStyle(sxTmp.SheetList[0].SheetName, sxTmp.SheetList[0].Title["地泵数据"][0]+fmt.Sprintf("%d", monDays+2), sxTmp.SheetList[0].Title["地泵数据"][0]+fmt.Sprintf("%d", monDays+2), style)
			}
		case "录像数据":
			if err = sxSumExcelFile.SetCellFormula(sxTmp.SheetList[0].SheetName, sxTmp.SheetList[0].Title["录像数据"][0]+fmt.Sprintf("%d", monDays+2), value); err != nil {
				fmt.Println(err)
				return
			} else {
				style, err = sxSumExcelFile.NewStyle(styleContent)
				if err != nil {
					fmt.Println(err)
					return
				}
				sxSumExcelFile.SetCellStyle(sxTmp.SheetList[0].SheetName, sxTmp.SheetList[0].Title["录像数据"][0]+fmt.Sprintf("%d", monDays+2), sxTmp.SheetList[0].Title["录像数据"][0]+fmt.Sprintf("%d", monDays+2), style)
			}

		}
	}
	// 填充剩余边框
	style, err = sxSumExcelFile.NewStyle(styleContent)
	if err != nil {
		fmt.Println(err)
		return
	}
	sxSumExcelFile.SetCellStyle(sxTmp.SheetList[0].SheetName, sxTmp.SheetList[0].Title["比对结果"][0]+fmt.Sprintf("%d", monDays+2), sxTmp.SheetList[0].Title["比对结果"][0]+fmt.Sprintf("%d", monDays+2), style)
	sxSumExcelFile.SetCellStyle(sxTmp.SheetList[0].SheetName, sxTmp.SheetList[0].Title["地泵明细"][0]+fmt.Sprintf("%d", monDays+2), sxTmp.SheetList[0].Title["地泵明细"][0]+fmt.Sprintf("%d", monDays+2), style)
	sxSumExcelFile.SetCellStyle(sxTmp.SheetList[0].SheetName, sxTmp.SheetList[0].Title["录像明细"][0]+fmt.Sprintf("%d", monDays+2), sxTmp.SheetList[0].Title["录像明细"][0]+fmt.Sprintf("%d", monDays+2), style)

	//---------------------------------------------------------------------------------------------------------------------------------
	// "tmpSheet"部分
	// 新建"tmpSheet"
	indexHz = sxSumExcelFile.NewSheet("tmp")
	sxSumExcelFile.SetActiveSheet(indexHz)

	// 设置"tmpSheet"页眉
	err = sxSumExcelFile.SetHeaderFooter("tmp", &excelize.FormatHeaderFooter{
		DifferentFirst: true,
		FirstHeader:    `&C` + `&B` + `&16` + `&"微软雅黑,常规"` + sxTmp.SheetList[0].Header,
	})
	if err != nil {
		fmt.Println(err)
		return
	}

	// 设置"tmpSheet"标题行
	// 设置"tmpSheet"一级标题
	if err = sxSumExcelFile.MergeCell(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["地泵数据"][0], sxTmp.SheetList[1].Title["地泵数据"][1]); err != nil {
		fmt.Println(err)
		return
	} else {
		if err = sxSumExcelFile.SetCellValue(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["地泵数据"][0], "地泵数据"); err != nil {
			fmt.Println(err)
			return
		} else {
			style, err = sxSumExcelFile.NewStyle(styleTitle)
			if err != nil {
				fmt.Println(err)
				return
			}
			sxSumExcelFile.SetCellStyle(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["地泵数据"][0], sxTmp.SheetList[1].Title["地泵数据"][1], style)
		}
	}
	if err = sxSumExcelFile.MergeCell(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["录像数据"][0], sxTmp.SheetList[1].Title["录像数据"][1]); err != nil {
		fmt.Println(err)
		return
	} else {
		if err = sxSumExcelFile.SetCellValue(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["录像数据"][0], "录像数据"); err != nil {
			fmt.Println(err)
			return
		} else {
			style, err = sxSumExcelFile.NewStyle(styleTitle)
			if err != nil {
				fmt.Println(err)
				return
			}
			sxSumExcelFile.SetCellStyle(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["录像数据"][0], sxTmp.SheetList[1].Title["录像数据"][1], style)
		}
	}

	// 设置"tmpSheet"二级标题
	for key, value := range sxTmp.SheetList[1].Title {
		// 设置"日期"列宽
		if err = sxSumExcelFile.SetColWidth(sxTmp.SheetList[1].SheetName, "A", "A", 12); err != nil {
			fmt.Println(err)
			return
		}

		if key != "地泵数据" && key != "录像数据" {
			if err = sxSumExcelFile.SetCellValue(sxTmp.SheetList[1].SheetName, value[1]+value[2], value[0]); err != nil {
				fmt.Println(err)
				return
			} else {
				style, err = sxSumExcelFile.NewStyle(styleTitle2)
				if err != nil {
					fmt.Println(err)
					return
				}
				sxSumExcelFile.SetCellStyle(sxTmp.SheetList[1].SheetName, value[1]+value[2], value[1]+value[2], style)
			}
		}
	}

	// 设置"tmpSheet"内容行

	for i := 1; i <= monDays; i++ {
		sxTmp.SheetList[1].ContentVariable["day"] = sxTmp.SheetList[1].ContentVariable["dayInit"] + fmt.Sprintf("%02d", i)

		// 日期列
		sxTmp.SheetList[1].ContentFixed["日期"] = sxTmp.SheetList[1].ContentVariable["day"]
		if err = sxSumExcelFile.SetCellValue(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["日期"][1]+fmt.Sprintf("%d", i+2), sxTmp.SheetList[1].ContentFixed["日期"]); err != nil {
			fmt.Println(err)
			return
		} else {
			style, err = sxSumExcelFile.NewStyle(styleContentAlignCenter)
			if err != nil {
				fmt.Println(err)
				return
			}
			sxSumExcelFile.SetCellStyle(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["日期"][1]+fmt.Sprintf("%d", i+2), sxTmp.SheetList[1].Title["日期"][1]+fmt.Sprintf("%d", i+2), style)
		}
		// 地泵总车数
		sxTmp.SheetList[1].ContentFixed["地泵总车数"] = "=VALUE(MID(汇总!$B" + fmt.Sprintf("%d", i+1) + ",FIND(\"总车数：\",汇总!$B" + fmt.Sprintf("%d", i+1) + ",1)+4,FIND(\" \",汇总!$B" + fmt.Sprintf("%d", i+1) + ",FIND(\"总车数：\",汇总!$B" + fmt.Sprintf("%d", i+1) + ",1)+4)-(FIND(\"总车数：\",汇总!$B" + fmt.Sprintf("%d", i+1) + ",1)+4)))"
		if err = sxSumExcelFile.SetCellFormula(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["地泵总车数"][1]+fmt.Sprintf("%d", i+2), sxTmp.SheetList[1].ContentFixed["地泵总车数"]); err != nil {
			fmt.Println(err)
			return
		} else {
			style, err = sxSumExcelFile.NewStyle(styleContentAlignCenter)
			if err != nil {
				fmt.Println(err)
				return
			}
			sxSumExcelFile.SetCellStyle(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["地泵总车数"][1]+fmt.Sprintf("%d", i+2), sxTmp.SheetList[1].Title["地泵总车数"][1]+fmt.Sprintf("%d", i+2), style)
		}
		// 地泵水渣
		sxTmp.SheetList[1].ContentFixed["地泵水渣"] = "=VALUE(MID(汇总!$B" + fmt.Sprintf("%d", i+1) + ",FIND(\"水渣：\",汇总!$B" + fmt.Sprintf("%d", i+1) + ",1)+3,FIND(\" \",汇总!$B" + fmt.Sprintf("%d", i+1) + ",FIND(\"水渣：\",汇总!$B" + fmt.Sprintf("%d", i+1) + ",1)+3)-(FIND(\"水渣：\",汇总!$B" + fmt.Sprintf("%d", i+1) + ",1)+3)))"
		if err = sxSumExcelFile.SetCellFormula(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["地泵水渣"][1]+fmt.Sprintf("%d", i+2), sxTmp.SheetList[1].ContentFixed["地泵水渣"]); err != nil {
			fmt.Println(err)
			return
		} else {
			style, err = sxSumExcelFile.NewStyle(styleContentAlignCenter)
			if err != nil {
				fmt.Println(err)
				return
			}
			sxSumExcelFile.SetCellStyle(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["地泵水渣"][1]+fmt.Sprintf("%d", i+2), sxTmp.SheetList[1].Title["地泵水渣"][1]+fmt.Sprintf("%d", i+2), style)
		}
		// 地泵矿粉
		sxTmp.SheetList[1].ContentFixed["地泵矿粉"] = "=VALUE(MID(汇总!$B" + fmt.Sprintf("%d", i+1) + ",FIND(\"矿粉：\",汇总!$B" + fmt.Sprintf("%d", i+1) + ",1)+3,FIND(\" \",汇总!$B" + fmt.Sprintf("%d", i+1) + ",FIND(\"矿粉：\",汇总!$B" + fmt.Sprintf("%d", i+1) + ",1)+3)-(FIND(\"矿粉：\",汇总!$B" + fmt.Sprintf("%d", i+1) + ",1)+3)))"
		if err = sxSumExcelFile.SetCellFormula(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["地泵矿粉"][1]+fmt.Sprintf("%d", i+2), sxTmp.SheetList[1].ContentFixed["地泵矿粉"]); err != nil {
			fmt.Println(err)
			return
		} else {
			style, err = sxSumExcelFile.NewStyle(styleContentAlignCenter)
			if err != nil {
				fmt.Println(err)
				return
			}
			sxSumExcelFile.SetCellStyle(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["地泵矿粉"][1]+fmt.Sprintf("%d", i+2), sxTmp.SheetList[1].Title["地泵矿粉"][1]+fmt.Sprintf("%d", i+2), style)
		}
		// 地泵其他
		sxTmp.SheetList[1].ContentFixed["地泵其他"] = "=VALUE(MID(汇总!$B" + fmt.Sprintf("%d", i+1) + ",FIND(\"其他：\",汇总!$B" + fmt.Sprintf("%d", i+1) + ",1)+3,3))"
		if err = sxSumExcelFile.SetCellFormula(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["地泵其他"][1]+fmt.Sprintf("%d", i+2), sxTmp.SheetList[1].ContentFixed["地泵其他"]); err != nil {
			fmt.Println(err)
			return
		} else {
			style, err = sxSumExcelFile.NewStyle(styleContentAlignCenter)
			if err != nil {
				fmt.Println(err)
				return
			}
			sxSumExcelFile.SetCellStyle(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["地泵其他"][1]+fmt.Sprintf("%d", i+2), sxTmp.SheetList[1].Title["地泵其他"][1]+fmt.Sprintf("%d", i+2), style)
		}
		// 录像总车数
		sxTmp.SheetList[1].ContentFixed["录像总车数"] = "=VALUE(MID(汇总!$C" + fmt.Sprintf("%d", i+1) + ",FIND(\"总车数：\",汇总!$C" + fmt.Sprintf("%d", i+1) + ",1)+4,FIND(\" \",汇总!$C" + fmt.Sprintf("%d", i+1) + ",FIND(\"总车数：\",汇总!$C" + fmt.Sprintf("%d", i+1) + ",1)+4)-(FIND(\"总车数：\",汇总!$C" + fmt.Sprintf("%d", i+1) + ",1)+4)))"
		if err = sxSumExcelFile.SetCellFormula(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["录像总车数"][1]+fmt.Sprintf("%d", i+2), sxTmp.SheetList[1].ContentFixed["录像总车数"]); err != nil {
			fmt.Println(err)
			return
		} else {
			style, err = sxSumExcelFile.NewStyle(styleContentAlignCenter)
			if err != nil {
				fmt.Println(err)
				return
			}
			sxSumExcelFile.SetCellStyle(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["录像总车数"][1]+fmt.Sprintf("%d", i+2), sxTmp.SheetList[1].Title["录像总车数"][1]+fmt.Sprintf("%d", i+2), style)
		}
		// 录像水渣
		sxTmp.SheetList[1].ContentFixed["录像水渣"] = "=VALUE(MID(汇总!$C" + fmt.Sprintf("%d", i+1) + ",FIND(\"水渣：\",汇总!$C" + fmt.Sprintf("%d", i+1) + ",1)+3,FIND(\" \",汇总!$C" + fmt.Sprintf("%d", i+1) + ",FIND(\"水渣：\",汇总!$C" + fmt.Sprintf("%d", i+1) + ",1)+3)-(FIND(\"水渣：\",汇总!$C" + fmt.Sprintf("%d", i+1) + ",1)+3)))"
		if err = sxSumExcelFile.SetCellFormula(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["录像水渣"][1]+fmt.Sprintf("%d", i+2), sxTmp.SheetList[1].ContentFixed["录像水渣"]); err != nil {
			fmt.Println(err)
			return
		} else {
			style, err = sxSumExcelFile.NewStyle(styleContentAlignCenter)
			if err != nil {
				fmt.Println(err)
				return
			}
			sxSumExcelFile.SetCellStyle(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["录像水渣"][1]+fmt.Sprintf("%d", i+2), sxTmp.SheetList[1].Title["录像水渣"][1]+fmt.Sprintf("%d", i+2), style)
		}
		// 录像矿粉
		sxTmp.SheetList[1].ContentFixed["录像矿粉"] = "=VALUE(MID(汇总!$C" + fmt.Sprintf("%d", i+1) + ",FIND(\"矿粉：\",汇总!$C" + fmt.Sprintf("%d", i+1) + ",1)+3,FIND(\" \",汇总!$C" + fmt.Sprintf("%d", i+1) + ",FIND(\"矿粉：\",汇总!$C" + fmt.Sprintf("%d", i+1) + ",1)+3)-(FIND(\"矿粉：\",汇总!$C" + fmt.Sprintf("%d", i+1) + ",1)+3)))"
		if err = sxSumExcelFile.SetCellFormula(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["录像矿粉"][1]+fmt.Sprintf("%d", i+2), sxTmp.SheetList[1].ContentFixed["录像矿粉"]); err != nil {
			fmt.Println(err)
			return
		} else {
			style, err = sxSumExcelFile.NewStyle(styleContentAlignCenter)
			if err != nil {
				fmt.Println(err)
				return
			}
			sxSumExcelFile.SetCellStyle(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["录像矿粉"][1]+fmt.Sprintf("%d", i+2), sxTmp.SheetList[1].Title["录像矿粉"][1]+fmt.Sprintf("%d", i+2), style)
		}
		// 录像其他
		sxTmp.SheetList[1].ContentFixed["录像其他"] = "=VALUE(MID(汇总!$C" + fmt.Sprintf("%d", i+1) + ",FIND(\"其他：\",汇总!$C" + fmt.Sprintf("%d", i+1) + ",1)+3,FIND(\" \",汇总!$C" + fmt.Sprintf("%d", i+1) + ",FIND(\"其他：\",汇总!$C" + fmt.Sprintf("%d", i+1) + ",1)+3)-(FIND(\"其他：\",汇总!$C" + fmt.Sprintf("%d", i+1) + ",1)+3)))"
		if err = sxSumExcelFile.SetCellFormula(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["录像其他"][1]+fmt.Sprintf("%d", i+2), sxTmp.SheetList[1].ContentFixed["录像其他"]); err != nil {
			fmt.Println(err)
			return
		} else {
			style, err = sxSumExcelFile.NewStyle(styleContentAlignCenter)
			if err != nil {
				fmt.Println(err)
				return
			}
			sxSumExcelFile.SetCellStyle(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["录像其他"][1]+fmt.Sprintf("%d", i+2), sxTmp.SheetList[1].Title["录像其他"][1]+fmt.Sprintf("%d", i+2), style)
		}
		// 录像异常
		sxTmp.SheetList[1].ContentFixed["录像异常"] = "=VALUE(MID(汇总!$C" + fmt.Sprintf("%d", i+1) + ",FIND(\"异常：\",汇总!$C" + fmt.Sprintf("%d", i+1) + ",1)+3,3))"
		if err = sxSumExcelFile.SetCellFormula(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["录像异常"][1]+fmt.Sprintf("%d", i+2), sxTmp.SheetList[1].ContentFixed["录像异常"]); err != nil {
			fmt.Println(err)
			return
		} else {
			style, err = sxSumExcelFile.NewStyle(styleContentAlignCenter)
			if err != nil {
				fmt.Println(err)
				return
			}
			sxSumExcelFile.SetCellStyle(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["录像异常"][1]+fmt.Sprintf("%d", i+2), sxTmp.SheetList[1].Title["录像异常"][1]+fmt.Sprintf("%d", i+2), style)
		}
	}
	// 设置"tmpSheet"汇总行
	// 汇总行"日期"
	if err = sxSumExcelFile.SetCellValue(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["日期"][1]+fmt.Sprintf("%d", monDays+3), "汇总"); err != nil {
		fmt.Println(err)
		return
	} else {
		style, err = sxSumExcelFile.NewStyle(styleTitle)
		if err != nil {
			fmt.Println(err)
			return
		}
		sxSumExcelFile.SetCellStyle(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["日期"][1]+fmt.Sprintf("%d", monDays+3), sxTmp.SheetList[1].Title["日期"][1]+fmt.Sprintf("%d", monDays+3), style)
	}
	// 地泵总车数汇总
	sxTmp.SheetList[1].Footer["地泵总车数汇总"] = "=SUM(B3:B" + fmt.Sprintf("%d", monDays+2) + ")"
	if err = sxSumExcelFile.SetCellFormula(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["地泵总车数"][1]+fmt.Sprintf("%d", monDays+3), sxTmp.SheetList[1].Footer["地泵总车数汇总"]); err != nil {
		fmt.Println(err)
		return
	} else {
		style, err = sxSumExcelFile.NewStyle(styleContentAlignCenter)
		if err != nil {
			fmt.Println(err)
			return
		}
		sxSumExcelFile.SetCellStyle(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["地泵总车数"][1]+fmt.Sprintf("%d", monDays+3), sxTmp.SheetList[1].Title["地泵总车数"][1]+fmt.Sprintf("%d", monDays+3), style)
	}
	// 地泵水渣汇总
	sxTmp.SheetList[1].Footer["地泵水渣汇总"] = "=SUM(C3:C" + fmt.Sprintf("%d", monDays+2) + ")"
	if err = sxSumExcelFile.SetCellFormula(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["地泵水渣"][1]+fmt.Sprintf("%d", monDays+3), sxTmp.SheetList[1].Footer["地泵水渣汇总"]); err != nil {
		fmt.Println(err)
		return
	} else {
		style, err = sxSumExcelFile.NewStyle(styleContentAlignCenter)
		if err != nil {
			fmt.Println(err)
			return
		}
		sxSumExcelFile.SetCellStyle(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["地泵水渣"][1]+fmt.Sprintf("%d", monDays+3), sxTmp.SheetList[1].Title["地泵水渣"][1]+fmt.Sprintf("%d", monDays+3), style)
	}
	// 地泵矿粉汇总
	sxTmp.SheetList[1].Footer["地泵矿粉汇总"] = "=SUM(D3:D" + fmt.Sprintf("%d", monDays+2) + ")"
	if err = sxSumExcelFile.SetCellFormula(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["地泵矿粉"][1]+fmt.Sprintf("%d", monDays+3), sxTmp.SheetList[1].Footer["地泵矿粉汇总"]); err != nil {
		fmt.Println(err)
		return
	} else {
		style, err = sxSumExcelFile.NewStyle(styleContentAlignCenter)
		if err != nil {
			fmt.Println(err)
			return
		}
		sxSumExcelFile.SetCellStyle(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["地泵矿粉"][1]+fmt.Sprintf("%d", monDays+3), sxTmp.SheetList[1].Title["地泵矿粉"][1]+fmt.Sprintf("%d", monDays+3), style)
	}
	// 地泵其他汇总
	sxTmp.SheetList[1].Footer["地泵其他汇总"] = "=SUM(E3:E" + fmt.Sprintf("%d", monDays+2) + ")"
	if err = sxSumExcelFile.SetCellFormula(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["地泵其他"][1]+fmt.Sprintf("%d", monDays+3), sxTmp.SheetList[1].Footer["地泵其他汇总"]); err != nil {
		fmt.Println(err)
		return
	} else {
		style, err = sxSumExcelFile.NewStyle(styleContentAlignCenter)
		if err != nil {
			fmt.Println(err)
			return
		}
		sxSumExcelFile.SetCellStyle(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["地泵其他"][1]+fmt.Sprintf("%d", monDays+3), sxTmp.SheetList[1].Title["地泵其他"][1]+fmt.Sprintf("%d", monDays+3), style)
	}
	// 录像总车数汇总
	sxTmp.SheetList[1].Footer["录像总车数汇总"] = "=SUM(G3:G" + fmt.Sprintf("%d", monDays+2) + ")"
	if err = sxSumExcelFile.SetCellFormula(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["录像总车数"][1]+fmt.Sprintf("%d", monDays+3), sxTmp.SheetList[1].Footer["录像总车数汇总"]); err != nil {
		fmt.Println(err)
		return
	} else {
		style, err = sxSumExcelFile.NewStyle(styleContentAlignCenter)
		if err != nil {
			fmt.Println(err)
			return
		}
		sxSumExcelFile.SetCellStyle(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["录像总车数"][1]+fmt.Sprintf("%d", monDays+3), sxTmp.SheetList[1].Title["录像总车数"][1]+fmt.Sprintf("%d", monDays+3), style)
	}
	// 录像水渣汇总
	sxTmp.SheetList[1].Footer["录像水渣汇总"] = "=SUM(H3:H" + fmt.Sprintf("%d", monDays+2) + ")"
	if err = sxSumExcelFile.SetCellFormula(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["录像水渣"][1]+fmt.Sprintf("%d", monDays+3), sxTmp.SheetList[1].Footer["录像水渣汇总"]); err != nil {
		fmt.Println(err)
		return
	} else {
		style, err = sxSumExcelFile.NewStyle(styleContentAlignCenter)
		if err != nil {
			fmt.Println(err)
			return
		}
		sxSumExcelFile.SetCellStyle(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["录像水渣"][1]+fmt.Sprintf("%d", monDays+3), sxTmp.SheetList[1].Title["录像水渣"][1]+fmt.Sprintf("%d", monDays+3), style)
	}
	// 录像矿粉汇总
	sxTmp.SheetList[1].Footer["录像矿粉汇总"] = "=SUM(I3:I" + fmt.Sprintf("%d", monDays+2) + ")"
	if err = sxSumExcelFile.SetCellFormula(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["录像矿粉"][1]+fmt.Sprintf("%d", monDays+3), sxTmp.SheetList[1].Footer["录像矿粉汇总"]); err != nil {
		fmt.Println(err)
		return
	} else {
		style, err = sxSumExcelFile.NewStyle(styleContentAlignCenter)
		if err != nil {
			fmt.Println(err)
			return
		}
		sxSumExcelFile.SetCellStyle(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["录像矿粉"][1]+fmt.Sprintf("%d", monDays+3), sxTmp.SheetList[1].Title["录像矿粉"][1]+fmt.Sprintf("%d", monDays+3), style)
	}
	// 录像其他汇总
	sxTmp.SheetList[1].Footer["录像其他汇总"] = "=SUM(J3:J" + fmt.Sprintf("%d", monDays+2) + ")"
	if err = sxSumExcelFile.SetCellFormula(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["录像其他"][1]+fmt.Sprintf("%d", monDays+3), sxTmp.SheetList[1].Footer["录像其他汇总"]); err != nil {
		fmt.Println(err)
		return
	} else {
		style, err = sxSumExcelFile.NewStyle(styleContentAlignCenter)
		if err != nil {
			fmt.Println(err)
			return
		}
		sxSumExcelFile.SetCellStyle(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["录像其他"][1]+fmt.Sprintf("%d", monDays+3), sxTmp.SheetList[1].Title["录像其他"][1]+fmt.Sprintf("%d", monDays+3), style)
	}
	// 录像异常汇总
	sxTmp.SheetList[1].Footer["录像异常汇总"] = "=SUM(K3:K" + fmt.Sprintf("%d", monDays+2) + ")"
	if err = sxSumExcelFile.SetCellFormula(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["录像异常"][1]+fmt.Sprintf("%d", monDays+3), sxTmp.SheetList[1].Footer["录像异常汇总"]); err != nil {
		fmt.Println(err)
		return
	} else {
		style, err = sxSumExcelFile.NewStyle(styleContentAlignCenter)
		if err != nil {
			fmt.Println(err)
			return
		}
		sxSumExcelFile.SetCellStyle(sxTmp.SheetList[1].SheetName, sxTmp.SheetList[1].Title["录像异常"][1]+fmt.Sprintf("%d", monDays+3), sxTmp.SheetList[1].Title["录像异常"][1]+fmt.Sprintf("%d", monDays+3), style)
	}
	return

}
