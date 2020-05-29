// Package api import "monitor/api"
package api

import (
	"fmt"
	"log"
	"os"
	"strconv"

	"github.com/zpbird/zp-go-mod/zdirfiles"
	"github.com/zpbird/zp-go-mod/zinput"
)

// 定义模板文件常量
const (
	sumTemplate   = "汇总模板.xlsx"
	videoTemplate = "录像明细模板.xlsx"
)

// GetCompany 需要传入模板文件夹的路径，返回选中的公司名称...
func GetCompany(templateDir string) string {

	// 获取模板文件夹下的公司名称列表
	// dirList := SubPathList(templateDir)
	dirList, err := zdirfiles.GetDirFileList(templateDir, "dir")
	if err != nil {
		log.Fatal(err)
		os.Exit(0)
	}
	var getNum int
	if len(dirList) == 0 {
		zinput.Clear()
		fmt.Println("模板目录：template 下没有任何文件夹，请按公司名称设置对应的文件夹！")
		os.Exit(0)
	}

	// 选择公司

	for trigger := true; trigger; {
		zinput.Clear()
		// 获取并显示公司列表
		tTermStr := "选择公司\n\n"
		for key, value := range dirList {
			if key+1 < 10 {
				tTermStr += " " + strconv.Itoa(key+1) + ". " + value + "\n"
			} else {
				tTermStr += strconv.Itoa(key+1) + ". " + value + "\n"
			}
		}
		tTermStr += "\n公司序号" + "[1 - " + strconv.Itoa(len(dirList)) + "] : "
		getStr := zinput.Input(tTermStr, zinput.RegNum)
		getNum, _ = strconv.Atoi(getStr)
		trigger = false

		if getNum < 1 || getNum > len(dirList) {
			zinput.StopContinue()
			trigger = true
		}
	}

	// 检查选中公司对应的模板文件是否存在
	if b1, _ := zdirfiles.DirFileExist(templateDir+"/"+dirList[getNum-1]+"/"+sumTemplate, "file"); !b1 {
		fmt.Println("该公司模板文件缺失！检查\"汇总.xlsx\"和\"录像明细.xlsx\"是否存在。")
		os.Exit(0)
	} else if b2, _ := zdirfiles.DirFileExist(templateDir+"/"+dirList[getNum-1]+"/"+videoTemplate, "file"); !b2 {
		fmt.Println("该公司模板文件缺失！检查\"汇总.xlsx\"和\"录像明细.xlsx\"是否存在。")
		os.Exit(0)
	}

	// 返回结果
	return dirList[getNum-1]
}
