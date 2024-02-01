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
	unknown, _ := com快捷类.I创建COM对象("Agent.Control.1")
	agent, _ := unknown.I查找接口(ole.IID_IDispatch)
	com快捷类.I设置属性(agent, "Connected", true)
	characters := com快捷类.I取属性P(agent, "Characters").ToIDispatch()
	com快捷类.I调用方法(characters, "Load", "Merlin", "c:\\windows\\msagent\\chars\\Merlin.acs")
	character := com快捷类.I调用方法P(characters, "Character", "Merlin").ToIDispatch()
	com快捷类.I调用方法(character, "Show")
	com快捷类.I调用方法(character, "Speak", "こんにちわ世界")

	time.Sleep(4000000000)
}
