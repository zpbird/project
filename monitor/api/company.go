// Package api import "monitor/api"
package api

import (
	"fmt"
	"os"
	"strconv"
)

// 定义模板文件常量
const (
	sumTemplate   = "汇总.xlsx"
	videoTemplate = "录像明细.xlsx"
)

// GetCompany 需要传入模板文件夹的路径，返回选中的公司名称...
func GetCompany(templateDir string) string {

	// 获取模板文件夹下的公司名称列表
	dirList := SubPathList(templateDir)
	var getNum int
	if len(dirList) == 0 {
		fmt.Println("模板文件夹：template 下没有任何文件夹，请按公司名称设置对应的文件夹！")
		os.Exit(0)
	}

	// 选择公司
	for getNum < 1 || getNum > len(dirList) {

		// 清屏
		Clear()

		// 获取并显示公司列表
		fmt.Printf("\n\n")
		fmt.Println("请选择公司(输入对应数字后回车)")
		fmt.Printf("\n")
		for key, value := range dirList {
			fmt.Printf(" %d. %s\n", key+1, value)
		}
		fmt.Printf("\n")
		fmt.Print("[1 - ", len(dirList), "] : ")

		// 验证输入的字符串为数字形式(一定程度上)
		var getStr string
		fmt.Scanln(&getStr)
		if i, err := strconv.ParseInt(getStr, 0, 64); err != nil {
			fmt.Println("输入时发生错误！")
			// os.Exit(0)
		} else {
			getNum = int(i)
		}
	}

	// 检查选中公司对应的模板文件是否存在
	if b1, _ := FileExists(templateDir + "/" + dirList[getNum-1] + "/" + sumTemplate); !b1 {
		fmt.Println("该公司模板文件缺失！检查\"汇总.xlsx\"和\"录像明细.xlsx\"是否存在。")
		os.Exit(0)
	} else if b2, _ := FileExists(templateDir + "/" + dirList[getNum-1] + "/" + videoTemplate); !b2 {
		fmt.Println("该公司模板文件缺失！检查\"汇总.xlsx\"和\"录像明细.xlsx\"是否存在。")
		os.Exit(0)
	}

	// 返回结果
	return dirList[getNum-1]
}
