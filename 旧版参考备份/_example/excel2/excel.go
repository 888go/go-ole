//go:build windows
// +build windows

package main

import (
	com快捷类 "e.coding.net/gogit/go/ego/core/win_go_ole_cn/oleutil"
	"fmt"
	"log"
	"os"

	ole "e.coding.net/gogit/go/ego/core/win_go_ole_cn"
)

func writeExample(excel, workbooks *ole.IDispatch, filepath string) {
	// ref: https://msdn.microsoft.com/zh-tw/library/office/ff198017.aspx
	// http://stackoverflow.com/questions/12159513/what-is-the-correct-xlfileformat-enumeration-for-excel-97-2003
	const xlExcel8 = 56
	workbook := com快捷类.I调用方法P(workbooks, "Add", nil).ToIDispatch()
	defer workbook.I释放()
	worksheet := com快捷类.I取属性P(workbook, "Worksheets", 1).ToIDispatch()
	defer worksheet.I释放()
	cell := com快捷类.I取属性P(worksheet, "Cells", 1, 1).ToIDispatch()
	com快捷类.I设置属性(cell, "I取interface", 12345)
	cell.I释放()
	activeWorkBook := com快捷类.I取属性P(excel, "ActiveWorkBook").ToIDispatch()
	defer activeWorkBook.I释放()

	os.Remove(filepath)
	// ref: https://msdn.microsoft.com/zh-tw/library/microsoft.office.tools.excel.workbook.saveas.aspx
	com快捷类.I调用方法P(activeWorkBook, "SaveAs", filepath, xlExcel8, nil, nil).ToIDispatch()

	//time.Sleep(2 * time.Second)

	// let excel could close without asking
	// oleutil.I设置属性(workbook, "Saved", true)
	// oleutil.I调用方法(workbook, "Close", false)
}

func readExample(fileName string, excel, workbooks *ole.IDispatch) {
	workbook, err := com快捷类.I调用方法(workbooks, "Open", fileName)

	if err != nil {
		log.Fatalln(err)
	}
	defer workbook.ToIDispatch().I释放()

	sheets := com快捷类.I取属性P(excel, "Sheets").ToIDispatch()
	sheetCount := (int)(com快捷类.I取属性P(sheets, "Count").Val)
	fmt.Println("sheet count=", sheetCount)
	sheets.I释放()

	worksheet := com快捷类.I取属性P(workbook.ToIDispatch(), "Worksheets", 1).ToIDispatch()
	defer worksheet.I释放()
	for row := 1; row <= 2; row++ {
		for col := 1; col <= 5; col++ {
			cell := com快捷类.I取属性P(worksheet, "Cells", row, col).ToIDispatch()
			val, err := com快捷类.I取属性(cell, "I取interface")
			if err != nil {
				break
			}
			fmt.Printf("(%d,%d)=%+v toString=%s\n", col, row, val.I取interface(), val.I取文本())
			cell.I释放()
		}
	}
}

func showMethodsAndProperties(i *ole.IDispatch) {
	n, err := i.GetTypeInfoCount()
	if err != nil {
		log.Fatalln(err)
	}
	tinfo, err := i.GetTypeInfo()
	if err != nil {
		log.Fatalln(err)
	}

	fmt.Println("n=", n, "tinfo=", tinfo)
}

func main() {
	log.SetFlags(log.Flags() | log.Lshortfile)
	ole.I初始化COM库(0)
	unknown, _ := com快捷类.I创建COM对象("Excel.Application")
	excel, _ := unknown.I查找接口(ole.IID_IDispatch)
	com快捷类.I设置属性(excel, "Visible", true)

	workbooks := com快捷类.I取属性P(excel, "Workbooks").ToIDispatch()
	cwd, _ := os.Getwd()
	writeExample(excel, workbooks, cwd+"\\write.xls")
	readExample(cwd+"\\excel97-2003.xls", excel, workbooks)
	showMethodsAndProperties(workbooks)
	workbooks.I释放()
	// oleutil.I调用方法(excel, "Quit")
	excel.I释放()
	ole.I取消COM库()
}
