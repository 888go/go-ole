//go:build windows
// +build windows

/*
	Demonstrates basic LibreOffce (OpenOffice) automation with OLE using GO-OLE.
	Usage: 	cd [...]\go-ole\example\libreoffice
			go run libreoffice.go
	References:
			http://www.openoffice.org/api/basic/man/tutorial/tutorial.pdf
			http://api.libreoffice.org/examples/examples.html#OLE_examples
			https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide

	Tested environment:
			go 1.6.2 (windows/amd64)
			LibreOffice 5.1.0.3 (32 bit)
			Windows 10 (64 bit)

	The MIT License (MIT)
	Copyright (c) 2016 Sebastian Schleemilch <https://github.com/itschleemilch>.

	Permission is hereby granted, free of charge, to any person obtaining a copy of
	this software and associated documentation files (the "Software"), to deal in
	the Software without restriction, including without limitation the rights to use,
	copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
	and to permit persons to whom the Software is furnished to do so, subject to the
	following conditions:

	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
	INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
	PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
	LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
	TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE
	OR OTHER DEALINGS IN THE SOFTWARE.
*/

package main

import (
	com快捷类 "e.coding.net/gogit/go/ego/core/win_go_ole_cn/oleutil"
	"fmt"
	"log"

	ole "e.coding.net/gogit/go/ego/core/win_go_ole_cn"
)

func checkError(err error, msg string) {
	if err != nil {
		log.Fatal(msg)
	}
}

// LOGetCell 返回工作表LibreOffice Basic中单元格的句柄: GetCell = oSheet.getCellByPosition (nColumn , nRow)
func LOGetCell(worksheet *ole.IDispatch, nColumn int, nRow int) (cell *ole.IDispatch) {
	return com快捷类.I调用方法P(worksheet, "getCellByPosition", nColumn, nRow).ToIDispatch()
}

// LOGetCellRangeByName 返回命名范围 (e.g. "A1:B4")
func LOGetCellRangeByName(worksheet *ole.IDispatch, rangeName string) (cells *ole.IDispatch) {
	return com快捷类.I调用方法P(worksheet, "getCellRangeByName", rangeName).ToIDispatch()
}

// LOGetCellString 返回显示的值
func LOGetCellString(cell *ole.IDispatch) (value string) {
	return com快捷类.I取属性P(cell, "string").I取文本()
}

// LOGetCellValue 返回单元格的内部值 (未格式化, 虚设代码, FIXME)
func LOGetCellValue(cell *ole.IDispatch) (value string) {
	val := com快捷类.I取属性P(cell, "value")
	fmt.Printf("Cell: %+v\n", val)
	return val.I取文本()
}

// LOGetCellError 返回单元格的错误值 (虚设代码, FIXME)
func LOGetCellError(cell *ole.IDispatch) (result *ole.I变体结构) {
	return com快捷类.I取属性P(cell, "error")
}

// LOSetCellString 设置单元格的文本值
func LOSetCellString(cell *ole.IDispatch, text string) {
	com快捷类.I设置属性P(cell, "string", text)
}

// LOSetCellValue 设置单元格的数值
func LOSetCellValue(cell *ole.IDispatch, value float64) {
	com快捷类.I设置属性P(cell, "value", value)
}

// LOSetCellFormula 设置公式（英语）
func LOSetCellFormula(cell *ole.IDispatch, formula string) {
	com快捷类.I设置属性P(cell, "formula", formula)
}

// LOSetCellFormulaLocal 以用户语言设置公式 (e.g. German =SUMME instead of =SUM)
func LOSetCellFormulaLocal(cell *ole.IDispatch, formula string) {
	com快捷类.I设置属性P(cell, "FormulaLocal", formula)
}

// LONewSpreadsheet 在新窗口中创建新的电子表格并返回文档句柄。
func LONewSpreadsheet(desktop *ole.IDispatch) (document *ole.IDispatch) {
	var args = []string{}
	document = com快捷类.I调用方法P(desktop,
		"loadComponentFromURL", "private:factory/scalc", // alternative: private:factory/swriter
		"_blank", 0, args).ToIDispatch()
	return
}

// LOOpenFile 打开文件 (text, spreadsheet, ...) 并返回文档句柄,
// 实例: /home/testuser/spreadsheet.ods
func LOOpenFile(desktop *ole.IDispatch, fullpath string) (document *ole.IDispatch) {
	var args = []string{}
	document = com快捷类.I调用方法P(desktop,
		"loadComponentFromURL", "file://"+fullpath,
		"_blank", 0, args).ToIDispatch()
	return
}

// LOSaveFile 保存当前文档。
// 仅当文件已经存在时，
// see https://wiki.openoffice.org/wiki/Saving_a_document
func LOSaveFile(document *ole.IDispatch) {
	// use storeAsURL if neccessary with third URL parameter
	com快捷类.I调用方法P(document, "store")
}

// LOGetWorksheet 返回工作表（索引从0开始）
func LOGetWorksheet(document *ole.IDispatch, index int) (worksheet *ole.IDispatch) {
	sheets := com快捷类.I取属性P(document, "Sheets").ToIDispatch()
	worksheet = com快捷类.I调用方法P(sheets, "getByIndex", index).ToIDispatch()
	return
}

// 本示例创建一个新的电子表格，读取并修改单元格值和样式。
func main() {
	ole.I初始化COM库(0)
	unknown, errCreate := com快捷类.I创建COM对象("com.sun.star.ServiceManager")
	checkError(errCreate, "Couldn't create a OLE connection to LibreOffice")
	ServiceManager, errSM := unknown.I查找接口(ole.IID_IDispatch)
	checkError(errSM, "Couldn't start a LibreOffice instance")
	desktop := com快捷类.I调用方法P(ServiceManager,
		"createInstance", "com.sun.star.frame.Desktop").ToIDispatch()

	document := LONewSpreadsheet(desktop)
	sheet0 := LOGetWorksheet(document, 0)

	cell1_1 := LOGetCell(sheet0, 1, 1) // cell B2
	cell1_2 := LOGetCell(sheet0, 1, 2) // cell B3
	cell1_3 := LOGetCell(sheet0, 1, 3) // cell B4
	cell1_4 := LOGetCell(sheet0, 1, 4) // cell B5
	LOSetCellString(cell1_1, "Hello World")
	LOSetCellValue(cell1_2, 33.45)
	LOSetCellFormula(cell1_3, "=B3+5")
	b4Value := LOGetCellString(cell1_3)
	LOSetCellString(cell1_4, b4Value)
	// set background color yellow:
	com快捷类.I设置属性P(cell1_1, "cellbackcolor", 0xFFFF00)

	fmt.Printf("Press [ENTER] to exit")
	fmt.Scanf("%s")
	ServiceManager.I释放()
	ole.I取消COM库()
}
