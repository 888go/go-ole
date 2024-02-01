package com子类

import "unsafe"

type IProvideClassInfo struct {
	IUnknown
}

type IProvideClassInfoVtbl struct {
	IUnknownVtbl
	GetClassInfo uintptr
}

func (v *IProvideClassInfo) VTable() *IProvideClassInfoVtbl {
	return (*IProvideClassInfoVtbl)(unsafe.Pointer(v.RawVTable))
}

// I取类信息 GetClassInfo
func (v *IProvideClassInfo) I取类信息() (cinfo *ITypeInfo, err error) {
	cinfo, err = getClassInfo(v)
	return
}
