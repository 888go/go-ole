package main

import (
	com子类 "e.coding.net/gogit/go/ego/core/win_go_ole_cn"
	com类 "e.coding.net/gogit/go/ego/core/win_go_ole_cn/oleutil"
	"fmt"
	//"github.com/go-ole/go-ole"
	//"github.com/go-ole/go-ole/oleutil"
)

func main() {
	//创建桌面快捷方式("C:\\1.exe", "E:\\Administrator\\Desktop\\123.lnk")
	com子类.I初始化COM库(0)
	defer com子类.I取消COM库()
	oleShellObject, err1 := com类.I创建COM对象("WScript.Shell")
	wshell, err2 := oleShellObject.I查找接口(com子类.IID_IDispatch)
	cs, err3 := com类.I调用方法(wshell, "CreateShortcut", "E:\\Administrator\\Desktop\\123.lnk")
	idispatch := cs.ToIDispatch()
	返回值2, _ := com类.I取属性(idispatch, "TargetPath")
	fmt.Println(返回值2.I取文本(), cs, err1, err2, err3)
}

//func 创建桌面快捷方式(src string, lnk名称 string) error {
//	//ole.CoInitialize(0)
//	//oleShellObject, err := oleutil.CreateObject("WScript.Shell")
//	//if err != nil {
//	//	return err
//	//}
//	//defer oleShellObject.Release()
//	//wshell, err := oleShellObject.QueryInterface(ole.IID_IDispatch)
//	//if err != nil {
//	//	return err
//	//}
//	//defer wshell.Release()
//	//cs, err := oleutil.CallMethod(wshell, "CreateShortcut", lnk名称)
//	//if err != nil {
//	//	return err
//	//}
//	//idispatch := cs.ToIDispatch()
//	//oleutil.PutProperty(idispatch, "TargetPath", src)
//	//oleutil.CallMethod(idispatch, "Save")
//	//return nil
//}
