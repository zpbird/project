// Package api import "monitor/api"
package api

import (
	"fmt"
	"os"
	"os/exec"
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
		// fmt.Printf("\x1bc") //清屏
		cmd := exec.Command("cmd.exe", "/c", "cls") //windows清屏命令
		cmd.Stdout = os.Stdout
		cmd.Run()

		fmt.Printf("\n\n")
		fmt.Println("请选择公司(输入对应数字)")
		fmt.Printf("\n")
		for key, value := range dirList {
			fmt.Printf(" %d. %s\n", key+1, value)
		}
		fmt.Printf("\n")
		fmt.Print("[1 - ", len(dirList), "] : ")

		fmt.Scanln(&getNum)
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
