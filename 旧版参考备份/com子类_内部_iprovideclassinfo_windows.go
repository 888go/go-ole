//go:build windows
// +build windows

package com子类

import (
	"syscall"
	"unsafe"
)

func getClassInfo(disp *IProvideClassInfo) (tinfo *ITypeInfo, err error) {
	hr, _, _ := syscall.Syscall(
		disp.VTable().GetClassInfo,
		2,
		uintptr(unsafe.Pointer(disp)),
		uintptr(unsafe.Pointer(&tinfo)),
		0)
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}
