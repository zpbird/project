// Package api import "monitor/api"
package api

import (
	"fmt"
	"os"
	"os/exec"
)

// GetCompany 需要传入模板文件夹的路径，返回选中的公司名称...
func GetCompany(templateDir string) string {
	// 获取模板文件夹下的公司名称列表
	if dirList := SubPathList(templateDir); len(dirList) == 0 {
		fmt.Println("模板文件夹：template 下没有任何文件夹，请按公司名称设置对应的文件夹！")
		os.Exit(0)
	} else {
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
		fmt.Println("( 1 -", len(dirList), ") : ")

	}

	// 选择公司

	// 检查选中公司对应的模板文件是否存在
	return "hello"
}
