package api

import (
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"os/exec"
	"sort"

	"github.com/Chain-Zhang/pinyin"
)

// PathExists 判断文件夹是否存在...
func PathExists(path string) (bool, error) {
	info, err := os.Stat(path)
	if err == nil && info.IsDir() {
		return true, nil
	}
	if os.IsNotExist(err) {
		return false, nil
	}
	return false, err
}

// FileExists 判断文件是否存在
func FileExists(file string) (bool, error) {
	info, err := os.Stat(file)
	if err == nil && !info.IsDir() {
		return true, nil
	}
	if os.IsNotExist(err) {
		return false, nil
	}
	return false, err
}

// SubPathList 获取子文件夹列表...
func SubPathList(path string) []string {

	files, err := ioutil.ReadDir(path)
	if err != nil {
		log.Fatal(err)
		os.Exit(0)
	}

	// 获取文件夹列表map
	dirMap := map[string]string{}
	for _, f := range files {
		if f.IsDir() {
			strPinyin, err := pinyin.New(f.Name()).Split("").Mode(pinyin.WithoutTone).Convert()
			if err != nil {
				fmt.Println("文件夹列表排序错误，退出！")
				os.Exit(0)
			} else {
				dirMap[strPinyin] = f.Name()
			}
		}
	}
	// fmt.Println(dirMap)

	// 生产按拼音排序后的slice

	tempSlice := []string{}
	dirSlice := []string{}

	for key := range dirMap {
		tempSlice = append(tempSlice, key)
	}
	sort.Strings(tempSlice)

	for _, key := range tempSlice {
		dirSlice = append(dirSlice, dirMap[key])
	}
	// fmt.Println(dirSlice)
	return dirSlice

}

// Clear 清屏...
func Clear() {
	// fmt.Printf("\x1bc") //清屏
	cmd := exec.Command("cmd.exe", "/c", "cls") //windows清屏命令
	cmd.Stdout = os.Stdout
	cmd.Run()
}
