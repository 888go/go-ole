//go:build windows
// +build windows

package com子类

import (
	"syscall"
	"unsafe"
)

// I取副本 Clone
func (enum *IEnumVARIANT) I取副本() (cloned *IEnumVARIANT, err error) {
	hr, _, _ := syscall.Syscall(
		enum.VTable().Clone,
		2,
		uintptr(unsafe.Pointer(enum)),
		uintptr(unsafe.Pointer(&cloned)),
		0)
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}

// I重置 Reset
func (enum *IEnumVARIANT) I重置() (err error) {
	hr, _, _ := syscall.Syscall(
		enum.VTable().Reset,
		1,
		uintptr(unsafe.Pointer(enum)),
		0,
		0)
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}

// I跳过 Skip
func (enum *IEnumVARIANT) I跳过(celt uint) (err error) {
	hr, _, _ := syscall.Syscall(
		enum.VTable().Skip,
		2,
		uintptr(unsafe.Pointer(enum)),
		uintptr(celt),
		0)
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}

// I下一个 Next
func (enum *IEnumVARIANT) I下一个(celt uint) (array I变体结构, length uint, err error) {
	hr, _, _ := syscall.Syscall6(
		enum.VTable().Next,
		4,
		uintptr(unsafe.Pointer(enum)),
		uintptr(celt),
		uintptr(unsafe.Pointer(&array)),
		uintptr(unsafe.Pointer(&length)),
		0,
		0)
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}
