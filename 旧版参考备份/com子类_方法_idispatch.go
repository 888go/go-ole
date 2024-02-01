package com子类

import "unsafe"

type IDispatch struct {
	IUnknown
}

type IDispatchVtbl struct {
	IUnknownVtbl
	GetTypeInfoCount uintptr
	GetTypeInfo      uintptr
	GetIDsOfNames    uintptr
	Invoke           uintptr
}

func (v *IDispatch) VTable() *IDispatchVtbl {
	return (*IDispatchVtbl)(unsafe.Pointer(v.RawVTable))
}

func (v *IDispatch) GetIDsOfName(names []string) (dispid []int32, err error) {
	dispid, err = getIDsOfName(v, names)
	return
}

func (v *IDispatch) Invoke(dispid int32, dispatch int16, params ...interface{}) (result *I变体结构, err error) {
	result, err = invoke(v, dispid, dispatch, params...)
	return
}

func (v *IDispatch) GetTypeInfoCount() (c uint32, err error) {
	c, err = getTypeInfoCount(v)
	return
}

func (v *IDispatch) GetTypeInfo() (tinfo *ITypeInfo, err error) {
	tinfo, err = getTypeInfo(v)
	return
}

// GetSingleIDOfName 是返回IDispatch名称的单个显示ID的助手。
//
// 这取代了试图从可用ID列表中获取单个名称的常见模式。它提供第一个ID（如果可用）。
func (v *IDispatch) GetSingleIDOfName(name string) (displayID int32, err error) {
	var displayIDs []int32
	displayIDs, err = v.GetIDsOfName([]string{name})
	if err != nil {
		return
	}
	displayID = displayIDs[0]
	return
}

// InvokeWithOptionalArgs 将参数作为数组接受，工作方式类似于Invoke。
//
// 接受名称并将尝试检索要传递给Invoke的显示ID。
//
// 将参数作为数组传递是一种解决方法，可以在以后的Go版本中修复，以防止传递空参数。
// 在测试过程中发现，这是一种可以接受的避免无法正常传递参数的方法。
func (v *IDispatch) InvokeWithOptionalArgs(name string, dispatch int16, params []interface{}) (result *I变体结构, err error) {
	displayID, err := v.GetSingleIDOfName(name)
	if err != nil {
		return
	}

	if len(params) < 1 {
		result, err = v.Invoke(displayID, dispatch)
	} else {
		result, err = v.Invoke(displayID, dispatch, params...)
	}

	return
}

// I调用方法  CallMethod 使用对象上的参数调用命名函数。//I调用方法
func (v *IDispatch) I调用方法(name string, params ...interface{}) (*I变体结构, error) {
	return v.InvokeWithOptionalArgs(name, DISPATCH_METHOD, params)
}

// I取属性 GetProperty 检索具有传递参数能力的名称的属性。
//
// 大多数情况下，您不需要传递参数，因为大多数对象不允许此功能。
// 或者至少不应允许此功能。有些服务器不遵循最佳实践，这是为那些边缘情况提供的。
func (v *IDispatch) I取属性(name string, params ...interface{}) (*I变体结构, error) {
	return v.InvokeWithOptionalArgs(name, DISPATCH_PROPERTYGET, params)
}

// I设置属性  PutProperty  改变对象中的属性。
func (v *IDispatch) I设置属性(name string, params ...interface{}) (*I变体结构, error) {
	return v.InvokeWithOptionalArgs(name, DISPATCH_PROPERTYPUT, params)
}
