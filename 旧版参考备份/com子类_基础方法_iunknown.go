package com子类

import "unsafe"

type IUnknown struct {
	RawVTable *interface{}
}

type IUnknownVtbl struct {
	QueryInterface uintptr
	AddRef         uintptr
	Release        uintptr
}

type UnknownLike interface {
	I查找接口(iid *GUID) (disp *IDispatch, err error)
	AddRef() int32
	I释放() int32
}

func (v *IUnknown) VTable() *IUnknownVtbl {
	return (*IUnknownVtbl)(unsafe.Pointer(v.RawVTable))
}

func (v *IUnknown) PutQueryInterface(interfaceID *GUID, obj interface{}) error {
	return reflectQueryInterface(v, v.VTable().QueryInterface, interfaceID, obj)
}

func (v *IUnknown) IDispatch(interfaceID *GUID) (dispatch *IDispatch, err error) {
	err = v.PutQueryInterface(interfaceID, &dispatch)
	return
}

func (v *IUnknown) IEnumVARIANT(interfaceID *GUID) (enum *IEnumVARIANT, err error) {
	err = v.PutQueryInterface(interfaceID, &enum)
	return
}

// I查找接口 QueryInterface
func (v *IUnknown) I查找接口(iid *GUID) (*IDispatch, error) {
	return queryInterface(v, iid)
}

// I查找接口p MustQueryInterface
func (v *IUnknown) I查找接口p(iid *GUID) (disp *IDispatch) {
	unk, err := queryInterface(v, iid)
	if err != nil {
		panic(err)
	}
	return unk
}

func (v *IUnknown) AddRef() int32 {
	return addRef(v)
}

// I释放 Release
func (v *IUnknown) I释放() int32 {
	return release(v)
}
