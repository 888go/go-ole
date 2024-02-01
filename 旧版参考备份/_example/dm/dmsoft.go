//go:build windows
// +build windows

// export GOARCH=386

package dmsoft

import (
	"syscall"
	"unsafe"

	ole "e.coding.net/gogit/go/ego/core/win_go_ole_cn"
	oleutil "e.coding.net/gogit/go/ego/core/win_go_ole_cn/oleutil"
)

var (
	dmReg32      = syscall.NewLazyDLL("DmReg.dll")
	_SetDllPathA = dmReg32.NewProc("SetDllPathA")
	_SetDllPathW = dmReg32.NewProc("SetDllPathW")
)

type Dmsoft struct {
	dm       *ole.IDispatch
	IUnknown *ole.IUnknown
}

func NewDmsoft() *Dmsoft {
	var com Dmsoft
	var err error
	ole.I初始化COM库(0)
	com.IUnknown, err = oleutil.I创建COM对象("dm.dmsoft")
	if err != nil {
		panic(err)
	}
	com.dm, err = com.IUnknown.I查找接口(ole.IID_IDispatch)
	if err != nil {
		panic(err)
	}
	return &com
}

// Release 释放
func (com *Dmsoft) Release() {
	com.IUnknown.I释放()
	com.dm.I释放()
	ole.I取消COM库()
}

// SetDllPathA Ascii
func SetDllPathA(path string, mode int) bool {
	var _p0 *uint16
	_p0, _ = syscall.UTF16PtrFromString(path)
	ret, _, _ := _SetDllPathA.Call(uintptr(unsafe.Pointer(_p0)), uintptr(mode))
	return ret != 0
}

// SetDllPathW Unicode
func SetDllPathW(path string, mode int) bool {
	var _p0 *uint16
	_p0, _ = syscall.UTF16PtrFromString(path)
	ret, _, _ := _SetDllPathW.Call(uintptr(unsafe.Pointer(_p0)), uintptr(mode))
	return ret != 0
}
