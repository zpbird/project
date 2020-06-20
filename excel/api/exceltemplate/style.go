package exceltemplate

import (
	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

var (
	sytleBorderAll = []excelize.Border{
		{Type: "top", Style: 1, Color: "000000"},
		{Type: "bottom", Style: 1, Color: "000000"},
		{Type: "left", Style: 1, Color: "000000"},
		{Type: "right", Style: 1, Color: "000000"},
	}
	styleAlignCenter = &excelize.Alignment{Vertical: "center", Horizontal: "center"}

	styleDefaultFont = "微软雅黑"
	styleTitle       = &excelize.Style{
		Font: &excelize.Font{
			Color: "000000", Bold: true, Size: 12, Family: "Microsoft YaHei"},
		Alignment: styleAlignCenter,
		Border:    sytleBorderAll,
	}
	styleTitle2 = &excelize.Style{
		Font: &excelize.Font{
			Color: "000000", Family: "Microsoft YaHei"},
		Alignment: styleAlignCenter,
		Border:    sytleBorderAll,
	}
	styleContentAlignCenter = &excelize.Style{
		Alignment: styleAlignCenter,
		Border:    sytleBorderAll,
	}
	styleContent = &excelize.Style{
		Border: sytleBorderAll,
	}
)
