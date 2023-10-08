package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"sort"
	"strings"
)

const (
	ImageSheet  = "Sheet1"
	OutputSheet = "Sheet3"
)

func main() {
	if len(os.Args) < 2 {
		panic("请输入文件路径，以及工作表名，例如 excel ./玉桂狗.xlsx")
	}
	excel, err := excelize.OpenFile(os.Args[1])
	if err != nil {
		panic(err)
	}

	// 获取列数和行数
	var colNum, lineNum int
	var BoundaryColorRGB string
	BoundaryColorRGB = "FFFFFF"
	for i := 1; ; i++ {
		if getCellBgColor(excel, ImageSheet, getColName(i)+"1") == BoundaryColorRGB {
			colNum = i - 1
			break
		}
	}
	for i := 1; ; i++ {
		if getCellBgColor(excel, ImageSheet, fmt.Sprintf("A%d", i)) == BoundaryColorRGB {
			lineNum = i - 1
			break
		}
	}
	fmt.Println("长高为:", colNum, lineNum)

	_ = excel.DeleteSheet(OutputSheet)
	_, err = excel.NewSheet(OutputSheet)
	if err != nil {
		panic(err)
	}

	//// 获取颜色对应的编号
	//// key RGB值，value 编码
	//colorMap := make(map[string]string)
	colorNum := make(map[string]int)
	//for i := 1; ; i++ {
	//	tmp := getCellBgColor(excel, OutputSheet, fmt.Sprintf("A%d", i))
	//	if tmp == BoundaryColorRGB {
	//		break
	//	}
	//	if value, ok := colorMap[tmp]; ok {
	//		panic(fmt.Sprintf("当前位置 A%d 的背景色编号已被定义为 %s，请重新编号", i, value))
	//	}
	//	value, err := excel.GetCellValue(OutputSheet, fmt.Sprintf("A%d", i))
	//	if err != nil {
	//		panic(err)
	//	}
	//	colorMap[tmp] = value
	//	colorNum[tmp] = 0
	//}

	// i代表当前进行的是第几行的遍历，由于 [A, lineNum] 处的格子只属于一行，所以总行数应该是colNum+lineNum-1
	for i := 1; i < colNum+lineNum; i++ {
		// 输出行号
		//fmt.Printf("%4d:", i)
		// 计算当前行的起始点的坐标
		//  横坐标			纵坐标
		var currentCol, currentLine int
		// 当前坐标颜色数量的计数器
		currentNum := 1
		if i < colNum {
			currentCol = colNum - i + 1
			currentLine = lineNum
		} else {
			currentCol = 1
			currentLine = colNum + lineNum - i
		}

		// 判断奇偶行，偶数行从左下向右上遍历
		// 			 奇数行从右上向左下遍历
		// 所以初始点位不同
		if i%2 != 0 {
			for {
				if currentCol == colNum || currentLine == 1 {
					break
				}
				currentLine--
				currentCol++
			}
		}
		// 获取当前单元格颜色
		outputColNum := 1
		outputLineNum := i
		currentColor := getCellBgColor(excel, ImageSheet, fmt.Sprintf("%s%d", getColName(currentCol), currentLine))
		for {
			// 判断奇偶行，偶数行从左下向右上遍历
			// 			 奇数行从右上向左下遍历
			// 直接进入下一个点位，初始化点位记录了
			if i%2 == 0 {
				currentLine--
				currentCol++
			} else {
				currentLine++
				currentCol--
			}
			if currentCol < 1 || currentLine < 1 || currentCol > colNum || currentLine > lineNum {
				// 本行结束
				//fmt.Printf("%d * %s\n", currentNum, colorMap[currentColor])
				cell := fmt.Sprintf("%s%d", getColName(outputColNum), outputLineNum)
				err = excel.SetCellValue(OutputSheet, cell, currentNum)
				if err != nil {
					panic(err)
				}
				err = setCellBgColor(excel, OutputSheet, cell, currentColor)
				if err != nil {
					panic(err)
				}
				colorNum[currentColor] += currentNum
				//color.HEXStyle("cca4e3", currentColor).Printf("  %2d  \n", currentNum)
				break
			}
			if currentColor == getCellBgColor(excel, ImageSheet, fmt.Sprintf("%s%d", getColName(currentCol), currentLine)) {
				currentNum++
			} else {
				// 颜色不同
				// 输出之前的颜色
				//fmt.Printf("%d * %s", currentNum, colorMap[currentColor])
				//color.HEXStyle("cca4e3", currentColor).Printf("  %2d  ", currentNum)
				cell := fmt.Sprintf("%s%d", getColName(outputColNum), outputLineNum)
				err = excel.SetCellValue(OutputSheet, cell, currentNum)
				if err != nil {
					panic(err)
				}
				err = setCellBgColor(excel, OutputSheet, cell, currentColor)
				if err != nil {
					panic(err)
				}
				outputColNum++
				colorNum[currentColor] += currentNum
				// 更新颜色信息
				currentColor = getCellBgColor(excel, ImageSheet, fmt.Sprintf("%s%d", getColName(currentCol), currentLine))
				currentNum = 1
			}
		}
	}
	// 输出每种颜色共多少个，输出到第八列，并按照数量进行排序
	var s []ColorNum
	for color, num := range colorNum {
		s = append(s, ColorNum{color, num})
	}
	sort.Slice(s, func(i, j int) bool {
		return s[i].num > s[j].num
	})
	for i, colorNum := range s {
		cell := fmt.Sprintf("%s%d", getColName(8), i+1)
		err = excel.SetCellValue(OutputSheet, cell, colorNum.num)
		if err != nil {
			panic(err)
		}
		err = setCellBgColor(excel, OutputSheet, cell, colorNum.color)
		if err != nil {
			panic(err)
		}
	}

	err = excel.Save()
	if err != nil {
		panic(err)
	}
}

type ColorNum struct {
	color string
	num   int
}

func setCellBgColor(f *excelize.File, sheet, cell, color string) error {
	style, err := f.NewStyle(&excelize.Style{
		Fill:      excelize.Fill{Type: "pattern", Color: []string{color}, Pattern: 1},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
	})
	if err != nil {
		return err
	}
	return f.SetCellStyle(sheet, cell, cell, style)
}

func getCellBgColor(f *excelize.File, sheet, cell string) string {
	styleID, err := f.GetCellStyle(sheet, cell)
	if err != nil {
		return err.Error()
	}
	fillID := *f.Styles.CellXfs.Xf[styleID].FillID
	fgColor := f.Styles.Fills.Fill[fillID].PatternFill.FgColor
	if fgColor != nil && f.Theme != nil {
		if clrScheme := f.Theme.ThemeElements.ClrScheme; fgColor.Theme != nil {
			if val, ok := map[int]*string{
				0: &clrScheme.Lt1.SysClr.LastClr,
				1: &clrScheme.Dk1.SysClr.LastClr,
				2: clrScheme.Lt2.SrgbClr.Val,
				3: clrScheme.Dk2.SrgbClr.Val,
				4: clrScheme.Accent1.SrgbClr.Val,
				5: clrScheme.Accent2.SrgbClr.Val,
				6: clrScheme.Accent3.SrgbClr.Val,
				7: clrScheme.Accent4.SrgbClr.Val,
				8: clrScheme.Accent5.SrgbClr.Val,
				9: clrScheme.Accent6.SrgbClr.Val,
			}[*fgColor.Theme]; ok && val != nil {
				return strings.TrimPrefix(excelize.ThemeColor(*val, fgColor.Tint), "FF")
			}
		}
		return strings.TrimPrefix(fgColor.RGB, "FF")
	}
	return "FFFFFF"
}

func getColName(index int) (colName string) {

	for index > 0 {
		remainder := (index - 1) % 26
		colName = string(rune('A'+remainder)) + colName
		index = (index - 1) / 26
	}
	return colName
}
