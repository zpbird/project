package main

import (
	"fmt"
	"monitor/api"
	"os"
	"strconv"

	"github.com/zpbird/zp-go-mod/zdirfiles"
	"github.com/zpbird/zp-go-mod/zinput"
	// "github.com/360EntSecGroup-Skylar/excelize/v2"
)

func main() {
	sysSep := string(os.PathSeparator)
	// 检查程序文件位置及模板主目录是否存在
	templateDir := api.CheckAppDir()

reInput: // 标签:重新输入

	// 选择公司
	selCompany := api.GetCompany(templateDir)

	// 选择时间段
	selYear, selSmon, selEmon := api.GetTime()

	// 确认选择信息
	for acceptSel := true; acceptSel; {
		monSection := ""
		if selEmon-selSmon == 0 {
			monSection = strconv.Itoa(selEmon)
		} else {
			monSection = strconv.Itoa(selSmon) + "-" + strconv.Itoa(selEmon)
		}

		zinput.Clear()
		tTermStr := "公司：" + selCompany + "     " + "年月份：" + strconv.Itoa(selYear) + "年" + monSection + "月份\n\n" + "y：确认  " + "n：重新选择\n\n" + "请输入选择[y 或 n]："
		if zinput.Input(tTermStr, zinput.RegYn) == "y" {
			acceptSel = false
		} else {
			goto reInput
		}

	}

	// fmt.Println(selCompany, selYear, selSmon, selEmon)

	// 创建相关目录
	targetRootDir := sysSep + "录像比对"
	targetCmpDir := targetRootDir + sysSep + selCompany
	targetDataDir := targetCmpDir + sysSep + "data"
	targetVideoDataDir := targetCmpDir + sysSep + "video_data"
	targetSumDir := targetCmpDir + sysSep + "汇总"

	if b, err := zdirfiles.MakeDir(targetRootDir); !b {
		fmt.Println(err)
		os.Exit(0)
	}
	if b, err := zdirfiles.MakeDir(targetCmpDir); !b {
		fmt.Println(err)
		os.Exit(0)
	}

	if b, err := zdirfiles.MakeDir(targetDataDir); !b {
		fmt.Println(err)
		os.Exit(0)
	}
	if b, err := zdirfiles.MakeDir(targetDataDir + sysSep + strconv.Itoa(selYear)); !b {
		fmt.Println(err)
		os.Exit(0)
	}

	if b, err := zdirfiles.MakeDir(targetVideoDataDir); !b {
		fmt.Println(err)
		os.Exit(0)
	}
	if b, err := zdirfiles.MakeDir(targetVideoDataDir + sysSep + strconv.Itoa(selYear)); !b {
		fmt.Println(err)
		os.Exit(0)
	}

	if b, err := zdirfiles.MakeDir(targetSumDir); !b {
		fmt.Println(err)
		os.Exit(0)
	}
	if b, err := zdirfiles.MakeDir(targetSumDir + sysSep + strconv.Itoa(selYear)); !b {
		fmt.Println(err)
		os.Exit(0)
	}

	for s, e := selSmon, selEmon; s <= e; s++ {
		if b, err := zdirfiles.MakeDir(targetDataDir + sysSep + strconv.Itoa(selYear) + sysSep + strconv.Itoa(s)); !b {
			fmt.Println(err)
			os.Exit(0)
		}
		if b, err := zdirfiles.MakeDir(targetVideoDataDir + sysSep + strconv.Itoa(selYear) + sysSep + strconv.Itoa(s)); !b {
			fmt.Println(err)
			os.Exit(0)
		}
	}

	// 拷贝录像明细文件

}
