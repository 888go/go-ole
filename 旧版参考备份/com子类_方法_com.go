//go:build windows
// +build windows

package com子类

import (
	"syscall"
	"unicode/utf16"
	"unsafe"
)

var (
	procCoInitialize            = modole32.NewProc("CoInitialize")
	procCoInitializeEx          = modole32.NewProc("CoInitializeEx")
	procCoUninitialize          = modole32.NewProc("CoUninitialize")
	procCoCreateInstance        = modole32.NewProc("CoCreateInstance")
	procCoTaskMemFree           = modole32.NewProc("CoTaskMemFree")
	procCLSIDFromProgID         = modole32.NewProc("CLSIDFromProgID")
	procCLSIDFromString         = modole32.NewProc("CLSIDFromString")
	procStringFromCLSID         = modole32.NewProc("StringFromCLSID")
	procStringFromIID           = modole32.NewProc("StringFromIID")
	procIIDFromString           = modole32.NewProc("IIDFromString")
	procCoGetObject             = modole32.NewProc("CoGetObject")
	procGetUserDefaultLCID      = modkernel32.NewProc("GetUserDefaultLCID")
	procCopyMemory              = modkernel32.NewProc("RtlMoveMemory")
	procVariantInit             = modoleaut32.NewProc("VariantInit")
	procVariantClear            = modoleaut32.NewProc("VariantClear")
	procVariantTimeToSystemTime = modoleaut32.NewProc("VariantTimeToSystemTime")
	procSysAllocString          = modoleaut32.NewProc("SysAllocString")
	procSysAllocStringLen       = modoleaut32.NewProc("SysAllocStringLen")
	procSysFreeString           = modoleaut32.NewProc("SysFreeString")
	procSysStringLen            = modoleaut32.NewProc("SysStringLen")
	procCreateDispTypeInfo      = modoleaut32.NewProc("CreateDispTypeInfo")
	procCreateStdDispatch       = modoleaut32.NewProc("CreateStdDispatch")
	procGetActiveObject         = modoleaut32.NewProc("GetActiveObject")

	procGetMessageW      = moduser32.NewProc("GetMessageW")
	procDispatchMessageW = moduser32.NewProc("DispatchMessageW")
)

// coInitialize 初始化当前线程上的COM库。
//
// MSDN文档建议不应调用此函数. 改为调用CoInitializeEx（）。
// 原因与线程有关，此功能仅适用于单线程公寓。
//
// 也就是说，该库的大多数用户都只使用了此功能。如果遇到线程问题，请使用  I初始化COM库_线程().
func coInitialize() (err error) {
	// http://msdn.microsoft.com/en-us/library/windows/desktop/ms678543(v=vs.85).aspx
	// 建议不应将任何值传递给 CoInitialized。
	// 由于参数是可选的，因此只能是Call（） <--需要测试才能确定。
	hr, _, _ := procCoInitialize.Call(uintptr(0))
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}

// coInitializeEx 使用并发模型初始化COM库。
func coInitializeEx(coinit uint32) (err error) {
	// http://msdn.microsoft.com/en-us/library/windows/desktop/ms695279(v=vs.85).aspx
	// 建议第一个参数不仅是可选的，而且应该始终为NULL。
	hr, _, _ := procCoInitializeEx.Call(uintptr(0), uintptr(coinit))
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}

// I初始化COM库 CoInitialize 初始化当前线程上的COM库。
//
// MSDN文档建议不应调用此函数. 改为调用CoInitializeEx（）。
// 原因与线程有关，此功能仅适用于单线程运行
//
// 也就是说，该库的大多数用户都只使用了此功能,
// 如果遇到线程问题，请使用
// I初始化COM库_线程().
func I初始化COM库(p uintptr) (err error) {
	// p 将被忽略，不会被使用。
	// 避免任何变量未使用的错误。
	p = uintptr(0)
	return coInitialize()
}

// I初始化COM库_线程  CoInitializeEx 使用并发模型初始化COM库
func I初始化COM库_线程(p uintptr, coinit uint32) (err error) {
	// p 将被忽略，不会被使用, 避免任何变量未使用的错误。
	p = uintptr(0)
	return coInitializeEx(coinit)
}

// I取消COM库  CoUninitialize 取消初始化 COM 库。
func I取消COM库() {
	procCoUninitialize.Call()
}

// I释放内存指针 CoTaskMemFree 释放内存指针。
func I释放内存指针(memptr uintptr) {
	procCoTaskMemFree.Call(memptr)
}

// I创建GUID对象_从程序ID  CLSIDFromProgID 检索具有给定程序标识符的类标识符。
//
// 必须注册编程标识符, 因为它将在Windows注册表中查找. 注册表项具有以 下项:
// 类id, Insertable, Protocol and Shell
// (https://msdn.microsoft.com/en-us/library/dd542719(v=vs.85).aspx).
//
// programID 以较低的精度标识类ID，并且不保证是唯一的.
// 这些通常在注册表中的 HKEY_LOCAL_MACHINE\SOFTWARE\Classes 下找到，
// 通常格式为  "Program.Component.Version" ，版本是可选的。
//
// I创建GUID对象_从程序ID CLSIDFromProgID 在 Windows API 中。
func I创建GUID对象_从程序ID(程序ID string) (guid对象 *GUID, err error) {
	var guid GUID
	lpszProgID := uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(程序ID)))
	hr, _, _ := procCLSIDFromProgID.Call(lpszProgID, uintptr(unsafe.Pointer(&guid)))
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	guid对象 = &guid
	return
}

// I创建GUID对象_从标识  CLSIDFromString 从字符串表示形式中检索类 ID。
//
// 从技术上讲，这是 GUID 的字符串版本，会将字符串转换为对象。
// 参数如: {248DD896-BB45-11CF-9ABC-0080C7E7B78D}
// I创建GUID对象_从标识 在 Windows API 中。
func I创建GUID对象_从标识(GUID标识 string) (guid对象 *GUID, err error) {
	var guid GUID
	lpsz := uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(GUID标识)))
	hr, _, _ := procCLSIDFromString.Call(lpsz, uintptr(unsafe.Pointer(&guid)))
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	guid对象 = &guid
	return
}

// GUID对象取标识  StringFromCLSID 从 GUID 对象返回 GUID 格式的字符串。
// 返回值如:{76A64158-CB41-11D1-8B02-00600806D9B6}
func GUID对象取标识(guid对象 *GUID) (str string, err error) {
	var p *uint16
	hr, _, _ := procStringFromCLSID.Call(uintptr(unsafe.Pointer(guid对象)), uintptr(unsafe.Pointer(&p)))
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	str = Unicode指针到文本(p)
	return
}

// I创建IID对象_从程序ID  IIDFromString 从程序ID 返回 GUID。
func I创建IID对象_从程序ID(progId string) (guid对象 *GUID, err error) {
	var guid GUID
	lpsz := uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(progId)))
	hr, _, _ := procIIDFromString.Call(lpsz, uintptr(unsafe.Pointer(&guid)))
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	guid对象 = &guid
	return
}

// IID对象取标识  StringFromIID 从 GUID 对象返回 GUID 格式的字符串。
// 返回值如: {76A64158-CB41-11D1-8B02-00600806D9B6}
func IID对象取标识(iid *GUID) (str string, err error) {
	var p *uint16
	hr, _, _ := procStringFromIID.Call(uintptr(unsafe.Pointer(iid)), uintptr(unsafe.Pointer(&p)))
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	str = Unicode指针到文本(p)
	return
}

// I创建COM对象 CreateInstance 具有 GUID 的单个未初始化对象。
// CoCreateInstance 函数创建指定的 CLSID 的一个实例，并返回客户端请求的类型的接口指针。
// 客户端负责通过在客户端使用完实例后调用其 Release 函数来管理实例的生存期。
// 若要基于单个 CLSID 创建多个对象，请调用 CoGetClassObject 函数。
// 若要连接到已创建并运行的对象，请调用 I取COM对象 函数。
func I创建COM对象(guid对象 *GUID, iid *GUID) (unk *IUnknown, err error) {
	if iid == nil {
		iid = IID_IUnknown
	}
	hr, _, _ := procCoCreateInstance.Call(
		uintptr(unsafe.Pointer(guid对象)),
		0,
		CLSCTX_SERVER,
		uintptr(unsafe.Pointer(iid)),
		uintptr(unsafe.Pointer(&unk)))
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}

// I取COM对象  GetActiveObject 检索已创建的IUnknown对象的指针。
// CoCreateInstance 函数创建指定的 CLSID 的一个实例，并返回客户端请求的类型的接口指针。
// 客户端负责通过在客户端使用完实例后调用其 Release 函数来管理实例的生存期。
// 若要基于单个 CLSID 创建多个对象，请调用 CoGetClassObject 函数。
// 若要连接到已创建并运行的对象，请调用 I取COM对象 函数。
func I取COM对象(guid对象 *GUID, iid *GUID) (unk *IUnknown, err error) {
	if iid == nil {
		iid = IID_IUnknown
	}
	hr, _, _ := procGetActiveObject.Call(
		uintptr(unsafe.Pointer(guid对象)),
		uintptr(unsafe.Pointer(iid)),
		uintptr(unsafe.Pointer(&unk)))
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}

type BindOpts struct {
	CbStruct          uint32
	GrfFlags          uint32
	GrfMode           uint32
	TickCountDeadline uint32
}

// GetObject 检索指向活动对象的指针。
func GetObject(programID string, bindOpts *BindOpts, iid *GUID) (unk *IUnknown, err error) {
	if bindOpts != nil {
		bindOpts.CbStruct = uint32(unsafe.Sizeof(BindOpts{}))
	}
	if iid == nil {
		iid = IID_IUnknown
	}
	hr, _, _ := procCoGetObject.Call(
		uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(programID))),
		uintptr(unsafe.Pointer(bindOpts)),
		uintptr(unsafe.Pointer(iid)),
		uintptr(unsafe.Pointer(&unk)))
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}

// I变体对象重置 VariantInit 初始化变体。
func I变体对象重置(v *I变体结构) (err error) {
	hr, _, _ := procVariantInit.Call(uintptr(unsafe.Pointer(v)))
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}

// I变体对象释放  VariantClear 以VT_EMPTY清除变量设置中的值。
// 使用此函数可以清除 VARIANTARG 类型变量 (或 I变体结构) ，然后释放包含 VARIANTARG 的内存 (，就像局部变量超出范围) 一样。
// 该函数通过将 vt 字段设置为VT_EMPTY清除 VARIANTARG。 VARIANTARG 的当前内容首先发布。 如果 vtfield VT_BSTR，则会释放字符串。 如果 vtfield VT_DISPATCH，则会释放该对象。 如果 vt 字段设置了VT_ARRAY位，则释放数组。
// 如果要清除的变体是引用传递的 COM 对象，则 pvargparameter 的 vtfield VT_DISPATCH |VT_BYREF或VT_UNKNOWN |VT_BYREF。 在这种情况下， VariantClear 不会释放对象。 由于被清除的变体是指向对对象的引用的指针， VariantClear 无法确定是否有必要释放对象。 因此，调用方有责任根据需要释放对象。
func I变体对象释放(v *I变体结构) (err error) {
	hr, _, _ := procVariantClear.Call(uintptr(unsafe.Pointer(v)))
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}

// SysAllocString 为字符串分配内存并将字符串复制到内存中。
func SysAllocString(v string) (ss *int16) {
	pss, _, _ := procSysAllocString.Call(uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(v))))
	ss = (*int16)(unsafe.Pointer(pss))
	return
}

// SysAllocStringLen 复制给定字符串返回指针的长度。
func SysAllocStringLen(v string) (ss *int16) {
	utf16 := utf16.Encode([]rune(v + "\x00"))
	ptr := &utf16[0]

	pss, _, _ := procSysAllocStringLen.Call(uintptr(unsafe.Pointer(ptr)), uintptr(len(utf16)-1))
	ss = (*int16)(unsafe.Pointer(pss))
	return
}

// SysFreeString 释放字符串系统内存。这必须使用 SysAllocString 调用。
func SysFreeString(v *int16) (err error) {
	hr, _, _ := procSysFreeString.Call(uintptr(unsafe.Pointer(v)))
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}

// SysStringLen 是系统分配的字符串的长度。
func SysStringLen(v *int16) uint32 {
	l, _, _ := procSysStringLen.Call(uintptr(unsafe.Pointer(v)))
	return uint32(l)
}

// CreateStdDispatch 为 IUnknown 提供默认的 IDispatch 实现。
//
// 这将处理对象的默认 IDispatch 实现. 它有一些限制，只支持一种语言。它也只会返回默认的异常代码。
func CreateStdDispatch(unk *IUnknown, v uintptr, ptinfo *IUnknown) (disp *IDispatch, err error) {
	hr, _, _ := procCreateStdDispatch.Call(
		uintptr(unsafe.Pointer(unk)),
		v,
		uintptr(unsafe.Pointer(ptinfo)),
		uintptr(unsafe.Pointer(&disp)))
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}

// CreateDispTypeInfo 为 IDispatch 提供默认的 ITypeInfo 实现。
//
// 这不会处理接口的完整实现。
func CreateDispTypeInfo(idata *INTERFACEDATA) (pptinfo *IUnknown, err error) {
	hr, _, _ := procCreateDispTypeInfo.Call(
		uintptr(unsafe.Pointer(idata)),
		uintptr(I取用户默认LCID()),
		uintptr(unsafe.Pointer(&pptinfo)))
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}

// copyMemory 移动内存块的位置。
func copyMemory(dest unsafe.Pointer, src unsafe.Pointer, length uint32) {
	procCopyMemory.Call(uintptr(dest), uintptr(src), uintptr(length))
}

// I取用户默认LCID 检索当前用户默认区域设置。
func I取用户默认LCID() (lcid uint32) {
	ret, _, _ := procGetUserDefaultLCID.Call()
	lcid = uint32(ret)
	return
}

// GetMessage 在运行时的消息队列中。
// 此功能似乎已阻止。速览消息不会阻止。
func GetMessage(msg *Msg, hwnd uint32, MsgFilterMin uint32, MsgFilterMax uint32) (ret int32, err error) {
	r0, _, err := procGetMessageW.Call(uintptr(unsafe.Pointer(msg)), uintptr(hwnd), uintptr(MsgFilterMin), uintptr(MsgFilterMax))
	ret = int32(r0)
	return
}

// DispatchMessage 到窗口程序。
func DispatchMessage(msg *Msg) (ret int32) {
	r0, _, _ := procDispatchMessageW.Call(uintptr(unsafe.Pointer(msg)))
	ret = int32(r0)
	return
}
