package main

//	"github.com/360EntSecGroup-Skylar/excelize/v2"
import (
	"fmt"
	"monitor/api"
	"os"
)

func main() {

	// 显示程序所在的目录
	sysSep := string(os.PathSeparator) // 当前系统路径分隔符
	pwd, err := os.Getwd()
	if err != nil {
		fmt.Println(err)
		// os.Exit(1)
	}
	fmt.Println("程序当前所在的目录为：", pwd)

	// 定义子目录变量
	// 模板目录
	templateDir := pwd + sysSep + "template"
	fmt.Println(templateDir)
	// 输出目录
	targetMidDir := pwd + sysSep + "output"
	fmt.Println(targetMidDir)
	// 选择公司名称
	cpnyList := api.GetCompanyList()
	fmt.Printf("%s\n", cpnyList)

	// 检查对应的模板文件是否存在

	// 选择时间段

}
