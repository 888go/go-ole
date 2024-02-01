package com子类

import (
	"fmt"
	"strings"
	"unsafe"
)

// DISPPARAMS 是传递给方法或属性的参数。
type DISPPARAMS struct {
	rgvarg            uintptr
	rgdispidNamedArgs uintptr
	cArgs             uint32
	cNamedArgs        uint32
}

// EXCEPINFO defines exception info.
type EXCEPINFO struct {
	wCode             uint16
	wReserved         uint16
	bstrSource        *uint16
	bstrDescription   *uint16
	bstrHelpFile      *uint16
	dwHelpContext     uint32
	pvReserved        uintptr
	pfnDeferredFillIn uintptr
	scode             uint32

	// 转到特定部分。不要移动上层，因为它会破坏本地代码的结构布局。
	rendered    bool
	source      string
	description string
	helpFile    string
}

// renderStrings 将BSTR字符串转换为Go ones，因此可以在“.I释放”之后安全地调用“.Error”和“.I取文本”。
// 当我们不能依靠调用者调用“.I释放”时，我们需要这样做。
func (e *EXCEPINFO) renderStrings() {
	e.rendered = true
	if e.bstrSource == nil {
		e.source = "<nil>"
	} else {
		e.source = BstrToString(e.bstrSource)
	}
	if e.bstrDescription == nil {
		e.description = "<nil>"
	} else {
		e.description = BstrToString(e.bstrDescription)
	}
	if e.bstrHelpFile == nil {
		e.helpFile = "<nil>"
	} else {
		e.helpFile = BstrToString(e.bstrHelpFile)
	}
}

// Clear 释放EXCEPINFO中的BSTR字符串并将其设置为NULL。
func (e *EXCEPINFO) Clear() {
	freeBSTR := func(s *uint16) {
		// SysFreeString don't return errors and is safe for call's on NULL.
		// https://docs.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-sysfreestring
		_ = SysFreeString((*int16)(unsafe.Pointer(s)))
	}

	if e.bstrSource != nil {
		freeBSTR(e.bstrSource)
		e.bstrSource = nil
	}
	if e.bstrDescription != nil {
		freeBSTR(e.bstrDescription)
		e.bstrDescription = nil
	}
	if e.bstrHelpFile != nil {
		freeBSTR(e.bstrHelpFile)
		e.bstrHelpFile = nil
	}
}

// WCode 返回EXCEPINFO中的wCode。
func (e EXCEPINFO) WCode() uint16 {
	return e.wCode
}

// SCODE 返回EXCEPINFO中的scode。
func (e EXCEPINFO) SCODE() uint32 {
	return e.scode
}

// I取文本 将EXCEPINFO转换为字符串。
func (e EXCEPINFO) I取文本() string {
	if !e.rendered {
		e.renderStrings()
	}
	return fmt.Sprintf(
		"wCode: %#x, bstrSource: %v, bstrDescription: %v, bstrHelpFile: %v, dwHelpContext: %#x, scode: %#x",
		e.wCode, e.source, e.description, e.helpFile, e.dwHelpContext, e.scode,
	)
}

// I取文本 将EXCEPINFO转换为字符串。
func (e EXCEPINFO) String() string {
	if !e.rendered {
		e.renderStrings()
	}
	return fmt.Sprintf(
		"wCode: %#x, bstrSource: %v, bstrDescription: %v, bstrHelpFile: %v, dwHelpContext: %#x, scode: %#x",
		e.wCode, e.source, e.description, e.helpFile, e.dwHelpContext, e.scode,
	)
}

// Error 实现错误接口并返回错误字符串。
func (e EXCEPINFO) Error() string {
	if !e.rendered {
		e.renderStrings()
	}

	if e.description != "<nil>" {
		return strings.TrimSpace(e.description)
	}

	code := e.scode
	if e.wCode != 0 {
		code = uint32(e.wCode)
	}
	return fmt.Sprintf("%v: %#x", e.source, code)
}

// PARAMDATA 定义参数数据类型。
type PARAMDATA struct {
	Name *int16
	Vt   uint16
}

// METHODDATA 定义方法信息。
type METHODDATA struct {
	Name     *uint16
	Data     *PARAMDATA
	Dispid   int32
	Meth     uint32
	CC       int32
	CArgs    uint32
	Flags    uint16
	VtReturn uint32
}

// INTERFACEDATA 定义接口信息。
type INTERFACEDATA struct {
	MethodData *METHODDATA
	CMembers   uint32
}

// Point is 2D vector type.
type Point struct {
	X int32
	Y int32
}

// Msg 是进程之间的消息。
type Msg struct {
	Hwnd    uint32
	Message uint32
	Wparam  int32
	Lparam  int32
	Time    uint32
	Pt      Point
}

// TYPEDESC 定义数据类型。
type TYPEDESC struct {
	Hreftype uint32
	VT       uint16
}

// IDLDESC 定义IDL信息。
type IDLDESC struct {
	DwReserved uint32
	WIDLFlags  uint16
}

// TYPEATTR 定义类型信息。
type TYPEATTR struct {
	Guid             GUID
	Lcid             uint32
	dwReserved       uint32
	MemidConstructor int32
	MemidDestructor  int32
	LpstrSchema      *uint16
	CbSizeInstance   uint32
	Typekind         int32
	CFuncs           uint16
	CVars            uint16
	CImplTypes       uint16
	CbSizeVft        uint16
	CbAlignment      uint16
	WTypeFlags       uint16
	WMajorVerNum     uint16
	WMinorVerNum     uint16
	TdescAlias       TYPEDESC
	IdldescType      IDLDESC
}
