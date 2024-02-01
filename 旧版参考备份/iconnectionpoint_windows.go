//go:build windows
// +build windows

package com子类

import (
	"syscall"
	"unsafe"
)

func (v *IConnectionPoint) GetConnectionInterface(piid **GUID) int32 {
	// XXX:这看起来不像是它应该做的
	return release((*IUnknown)(unsafe.Pointer(v)))
}

func (v *IConnectionPoint) Advise(unknown *IUnknown) (cookie uint32, err error) {
	hr, _, _ := syscall.Syscall(
		v.VTable().Advise,
		3,
		uintptr(unsafe.Pointer(v)),
		uintptr(unsafe.Pointer(unknown)),
		uintptr(unsafe.Pointer(&cookie)))
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}

func (v *IConnectionPoint) Unadvise(cookie uint32) (err error) {
	hr, _, _ := syscall.Syscall(
		v.VTable().Unadvise,
		2,
		uintptr(unsafe.Pointer(v)),
		uintptr(cookie),
		0)
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}

func (v *IConnectionPoint) EnumConnections(p *unsafe.Pointer) error {
	return I创建错误对象(E_NOTIMPL)
}
