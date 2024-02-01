//go:build windows
// +build windows

package main

import (
	com快捷类 "e.coding.net/gogit/go/ego/core/win_go_ole_cn/oleutil"
	"fmt"
	"log"

	ole "e.coding.net/gogit/go/ego/core/win_go_ole_cn"
)

func main() {
	ole.I初始化COM库(0)
	unknown, err := com快捷类.I创建COM对象("WMPlayer.OCX")
	if err != nil {
		log.Fatal(err)
	}
	wmp := unknown.I查找接口p(ole.IID_IDispatch)
	collection := com快捷类.I取属性P(wmp, "MediaCollection").ToIDispatch()
	list := com快捷类.I调用方法P(collection, "getAll").ToIDispatch()
	count := int(com快捷类.I取属性P(list, "count").Val)
	for i := 0; i < count; i++ {
		item := com快捷类.I取属性P(list, "item", i).ToIDispatch()
		name := com快捷类.I取属性P(item, "name").I取文本()
		sourceURL := com快捷类.I取属性P(item, "sourceURL").I取文本()
		fmt.Println(name, sourceURL)
	}
}
