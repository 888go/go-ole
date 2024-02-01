//go:build windows
// +build windows

package main

import (
	com快捷类 "e.coding.net/gogit/go/ego/core/win_go_ole_cn/oleutil"
	"fmt"

	ole "e.coding.net/gogit/go/ego/core/win_go_ole_cn"
)

func main() {
	ole.I初始化COM库(0)
	unknown, _ := com快捷类.I创建COM对象("Outlook.Application")
	outlook, _ := unknown.I查找接口(ole.IID_IDispatch)
	ns := com快捷类.I调用方法P(outlook, "GetNamespace", "MAPI").ToIDispatch()
	folder := com快捷类.I调用方法P(ns, "GetDefaultFolder", 10).ToIDispatch()
	contacts := com快捷类.I调用方法P(folder, "Items").ToIDispatch()
	count := com快捷类.I取属性P(contacts, "Count").I取interface().(int32)
	for i := 1; i <= int(count); i++ {
		item, err := com快捷类.I取属性(contacts, "Item", i)
		if err == nil && item.VT == ole.VT_DISPATCH {
			if value, err := com快捷类.I取属性(item.ToIDispatch(), "FullName"); err == nil {
				fmt.Println(value.I取interface())
			}
		}
	}
	com快捷类.I调用方法P(outlook, "Quit")
}
