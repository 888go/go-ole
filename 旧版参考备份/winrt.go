//go:build windows
// +build windows

package com子类

import (
	"reflect"
	"syscall"
	"unicode/utf8"
	"unsafe"
)

var (
	procRoInitialize              = modcombase.NewProc("RoInitialize")
	procRoActivateInstance        = modcombase.NewProc("RoActivateInstance")
	procRoGetActivationFactory    = modcombase.NewProc("RoGetActivationFactory")
	procWindowsCreateString       = modcombase.NewProc("WindowsCreateString")
	procWindowsDeleteString       = modcombase.NewProc("WindowsDeleteString")
	procWindowsGetStringRawBuffer = modcombase.NewProc("WindowsGetStringRawBuffer")
)

func RoInitialize(thread_type uint32) (err error) {
	hr, _, _ := procRoInitialize.Call(uintptr(thread_type))
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}

func RoActivateInstance(guid对象 string) (ins *IInspectable, err error) {
	hClsid, err := I创建HString对象(guid对象)
	if err != nil {
		return nil, err
	}
	defer I删除HString对象(hClsid)

	hr, _, _ := procRoActivateInstance.Call(
		uintptr(unsafe.Pointer(hClsid)),
		uintptr(unsafe.Pointer(&ins)))
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}

func RoGetActivationFactory(guid对象 string, iid *GUID) (ins *IInspectable, err error) {
	hClsid, err := I创建HString对象(guid对象)
	if err != nil {
		return nil, err
	}
	defer I删除HString对象(hClsid)

	hr, _, _ := procRoGetActivationFactory.Call(
		uintptr(unsafe.Pointer(hClsid)),
		uintptr(unsafe.Pointer(iid)),
		uintptr(unsafe.Pointer(&ins)))
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}

// HString 是指针的句柄字符串。
type HString uintptr

// I创建HString对象 NewHString 返回Go字符串的新HString。
func I创建HString对象(s string) (hstring HString, err error) {
	u16 := syscall.StringToUTF16Ptr(s)
	len := uint32(utf8.RuneCountInString(s))
	hr, _, _ := procWindowsCreateString.Call(
		uintptr(unsafe.Pointer(u16)),
		uintptr(len),
		uintptr(unsafe.Pointer(&hstring)))
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}

// I删除HString对象  DeleteHString 删除HString。
func I删除HString对象(hstring HString) (err error) {
	hr, _, _ := procWindowsDeleteString.Call(uintptr(hstring))
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}

// I取文本  String 返回HString的Go字符串值。
func (h HString) I取文本() string {
	var u16buf uintptr
	var u16len uint32
	u16buf, _, _ = procWindowsGetStringRawBuffer.Call(
		uintptr(h),
		uintptr(unsafe.Pointer(&u16len)))

	u16hdr := reflect.SliceHeader{Data: u16buf, Len: int(u16len), Cap: int(u16len)}
	u16 := *(*[]uint16)(unsafe.Pointer(&u16hdr))
	return syscall.UTF16ToString(u16)
}

// I取文本 返回HString的Go字符串值。
func (h HString) String() string {
	var u16buf uintptr
	var u16len uint32
	u16buf, _, _ = procWindowsGetStringRawBuffer.Call(
		uintptr(h),
		uintptr(unsafe.Pointer(&u16len)))

	u16hdr := reflect.SliceHeader{Data: u16buf, Len: int(u16len), Cap: int(u16len)}
	u16 := *(*[]uint16)(unsafe.Pointer(&u16hdr))
	return syscall.UTF16ToString(u16)
}
