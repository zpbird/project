package api

import (
	"fmt"
	"os"
	"strings"
)

// CheckAppDir 检查程序基本目录情况...
func CheckAppDir() string {
	// 当前系统路径分隔符
	sysSep := string(os.PathSeparator)

	// 程序所在目录
	pwd, err := os.Getwd()
	if err != nil {
		fmt.Println(err)
		os.Exit(0)
	}
	if pwd == pwd[0:strings.Index(pwd, sysSep)+1] {
		fmt.Println("当前程序处于磁盘根目录", pwd, "，请将相关文件放入某个文件夹中！")
		os.Exit(0)
	}

	// 模板目录
	templateDir := pwd + sysSep + "template"
	if ex, err := PathExists(templateDir); !ex || err != nil {
		fmt.Println("模板文件夹：template 不存在！")
		os.Exit(0)
	}

	return templateDir
}
