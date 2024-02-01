//go:build windows
// +build windows

package main

import (
	com快捷类 "e.coding.net/gogit/go/ego/core/win_go_ole_cn/oleutil"
	"fmt"
	"time"

	ole "e.coding.net/gogit/go/ego/core/win_go_ole_cn"
)

func main() {
	ole.I初始化COM库(0)
	unknown, _ := com快捷类.I创建COM对象("Microsoft.XMLHTTP")
	xmlhttp, _ := unknown.I查找接口(ole.IID_IDispatch)
	_, err := com快捷类.I调用方法(xmlhttp, "open", "GET", "http://rss.slashdot.org/Slashdot/slashdot", false)
	if err != nil {
		panic(err.Error())
	}
	_, err = com快捷类.I调用方法(xmlhttp, "send", nil)
	if err != nil {
		panic(err.Error())
	}
	state := -1
	for state != 4 {
		state = int(com快捷类.I取属性P(xmlhttp, "readyState").Val)
		time.Sleep(10000000)
	}
	responseXml := com快捷类.I取属性P(xmlhttp, "responseXml").ToIDispatch()
	items := com快捷类.I调用方法P(responseXml, "selectNodes", "/rdf:RDF/item").ToIDispatch()
	length := int(com快捷类.I取属性P(items, "length").Val)

	println(length)
	for n := 0; n < length; n++ {
		item := com快捷类.I取属性P(items, "item", n).ToIDispatch()

		title := com快捷类.I调用方法P(item, "selectSingleNode", "title").ToIDispatch()
		fmt.Println(com快捷类.I取属性P(title, "text").I取文本())

		link := com快捷类.I调用方法P(item, "selectSingleNode", "link").ToIDispatch()
		fmt.Println("  " + com快捷类.I取属性P(link, "text").I取文本())

		title.I释放()
		link.I释放()
		item.I释放()
	}
	items.I释放()
	xmlhttp.I释放()
}
