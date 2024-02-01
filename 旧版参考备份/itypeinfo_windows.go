//go:build windows
// +build windows

package com子类

import (
	"syscall"
	"unsafe"
)

// I取类型属性 GetTypeAttr
func (v *ITypeInfo) I取类型属性() (tattr *TYPEATTR, err error) {
	hr, _, _ := syscall.Syscall(
		uintptr(v.VTable().GetTypeAttr),
		2,
		uintptr(unsafe.Pointer(v)),
		uintptr(unsafe.Pointer(&tattr)),
		0)
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}
