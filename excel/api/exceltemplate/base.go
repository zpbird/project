// Package exceltemplate import "monitor/api/exceltemplate" or "excel/api/exceltemplate"
package exceltemplate

import (
	"os"
)

var sysSep = string(os.PathSeparator)

// SumExcelTmp 汇总表结构体
type SumExcelTmp struct {
	SheetList []struct {
		SheetName       string
		Header          string
		Title           map[string][2]string
		ContentVariable map[string]string
		ContentFixed    map[string]string
		Footer          map[string]string
	}
}

// NewSumExcelTmp ...
func NewSumExcelTmp() *SumExcelTmp {
	return &SumExcelTmp{
		SheetList: make([]struct {
			SheetName       string
			Header          string
			Title           map[string][2]string
			ContentVariable map[string]string
			ContentFixed    map[string]string
			Footer          map[string]string
		}, 3), // 是否应该默认为0，或指定个数
	}
}
