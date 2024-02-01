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
	unknown, _ := com快捷类.I创建COM对象("InternetExplorer.Application")
	ie, _ := unknown.I查找接口(ole.IID_IDispatch)
	com快捷类.I设置属性(ie, "Visible", true)
	com快捷类.I调用方法(ie, "Navigate", "http://www.google.com")
	for {
		if com快捷类.I取属性P(ie, "Busy").Val == 0 {
			break
		}
	}

	time.Sleep(1e9)

	document := com快捷类.I取属性P(ie, "document").ToIDispatch()

	// 将“golang”设置为文本框。
	elems := com快捷类.I调用方法P(document, "getElementsByName", "q").ToIDispatch()
	q := com快捷类.I调用方法P(elems, "item", 0).ToIDispatch()
	com快捷类.I设置属性P(q, "value", "golang")

	// 单击btnK。
	elems = com快捷类.I调用方法P(document, "getElementsByName", "btnK").ToIDispatch()
	btnG := com快捷类.I调用方法P(elems, "item", 0).ToIDispatch()
	com快捷类.I调用方法P(btnG, "click")
}
