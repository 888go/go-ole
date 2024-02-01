//go:build windows
// +build windows

package main

import (
	com快捷类 "e.coding.net/gogit/go/ego/core/win_go_ole_cn/oleutil"
	"time"

	ole "e.coding.net/gogit/go/ego/core/win_go_ole_cn"
)

func main() {
	ole.I初始化COM库(0)
	unknown, _ := com快捷类.I创建COM对象("Excel.Application")
	excel, _ := unknown.I查找接口(ole.IID_IDispatch)
	com快捷类.I设置属性(excel, "Visible", true)
	workbooks := com快捷类.I取属性P(excel, "Workbooks").ToIDispatch()
	workbook := com快捷类.I调用方法P(workbooks, "Add", nil).ToIDispatch()
	worksheet := com快捷类.I取属性P(workbook, "Worksheets", 1).ToIDispatch()
	cell := com快捷类.I取属性P(worksheet, "Cells", 1, 1).ToIDispatch()
	com快捷类.I设置属性(cell, "I取interface", 12345)

	time.Sleep(2000000000)

	com快捷类.I设置属性(workbook, "Saved", true)
	com快捷类.I调用方法(workbook, "Close", false)
	com快捷类.I调用方法(excel, "Quit")
	excel.I释放()

	ole.I取消COM库()
}
