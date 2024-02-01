//go:build windows
// +build windows

package com子类

import (
	"syscall"
	"unsafe"
)

func (v *IConnectionPointContainer) EnumConnectionPoints(points interface{}) error {
	return I创建错误对象(E_NOTIMPL)
}

func (v *IConnectionPointContainer) FindConnectionPoint(iid *GUID, point **IConnectionPoint) (err error) {
	hr, _, _ := syscall.Syscall(
		v.VTable().FindConnectionPoint,
		3,
		uintptr(unsafe.Pointer(v)),
		uintptr(unsafe.Pointer(iid)),
		uintptr(unsafe.Pointer(point)))
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}
