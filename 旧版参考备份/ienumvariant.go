package com子类

import "unsafe"

type IEnumVARIANT struct {
	IUnknown
}

type IEnumVARIANTVtbl struct {
	IUnknownVtbl
	Next  uintptr
	Skip  uintptr
	Reset uintptr
	Clone uintptr
}

func (v *IEnumVARIANT) VTable() *IEnumVARIANTVtbl {
	return (*IEnumVARIANTVtbl)(unsafe.Pointer(v.RawVTable))
}
