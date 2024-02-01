package com类

import com类 "e.coding.net/gogit/go/ego/core/win_go_ole_cn"

// I创建GUID对象  ClassIDFrom 检索类ID，无论给定的是程序ID还是应用程序字符串。
func I创建GUID对象(programID string) (guid对象 *com类.GUID, err error) {
	return com类.I创建GUID对象(programID)
}

// I创建COM对象 CreateObject 根据接口类型从programID创建对象。
//
// 仅支持IUnknown。
//
// 程序ID可以是程序ID或应用程序字符串。
func I创建COM对象(programID string) (COM对象 *com类.IUnknown, err error) {
	classID, err := com类.I创建GUID对象(programID)
	if err != nil {
		return
	}

	COM对象, err = com类.I创建COM对象(classID, com类.IID_IUnknown)
	if err != nil {
		return
	}

	return
}

// I取COM对象  GetActiveObject 根据接口类型检索程序ID和接口ID的活动对象。
//
// 仅支持IUnknown。
//
// 程序ID可以是程序ID或应用程序字符串。
func I取COM对象(programID string) (COM对象 *com类.IUnknown, err error) {
	classID, err := com类.I创建GUID对象(programID)
	if err != nil {
		return
	}

	COM对象, err = com类.I取COM对象(classID, com类.IID_IUnknown)
	if err != nil {
		return
	}

	return
}

// I调用方法  CallMethod 使用参数调用IDispatch上的方法。
func I调用方法(disp *com类.IDispatch, 方法名 string, 参数s ...interface{}) (result *com类.I变体结构, err error) {
	return disp.InvokeWithOptionalArgs(方法名, com类.DISPATCH_METHOD, 参数s)
}

// I调用方法P  MustCallMethod 使用参数或panic调用IDispatch上的方法。
func I调用方法P(disp *com类.IDispatch, 方法名 string, 参数s ...interface{}) (result *com类.I变体结构) {
	r, err := I调用方法(disp, 方法名, 参数s...)
	if err != nil {
		panic(err.Error())
	}
	return r
}

// I取属性 GetProperty 从IDispatch中检索属性。
func I取属性(disp *com类.IDispatch, 属性名 string, 参数s ...interface{}) (result *com类.I变体结构, err error) {
	return disp.InvokeWithOptionalArgs(属性名, com类.DISPATCH_PROPERTYGET, 参数s)
}

// I取属性P  MustGetProperty 从IDispatch或panic中检索属性。
func I取属性P(disp *com类.IDispatch, 属性名 string, 参数s ...interface{}) (result *com类.I变体结构) {
	r, err := I取属性(disp, 属性名, 参数s...)
	if err != nil {
		panic(err.Error())
	}
	return r
}

// I设置属性  PutProperty 改变对象中的属性
func I设置属性(disp *com类.IDispatch, 属性名 string, 参数s ...interface{}) (result *com类.I变体结构, err error) {
	return disp.InvokeWithOptionalArgs(属性名, com类.DISPATCH_PROPERTYPUT, 参数s)
}

// I设置属性P MustPutProperty 改变对象中的属性, 失败会恐慌。
func I设置属性P(disp *com类.IDispatch, 属性名 string, 参数s ...interface{}) (result *com类.I变体结构) {
	r, err := I设置属性(disp, 属性名, 参数s...)
	if err != nil {
		panic(err.Error())
	}
	return r
}

// PutPropertyRef 变异属性引用。
func PutPropertyRef(disp *com类.IDispatch, name string, 参数s ...interface{}) (result *com类.I变体结构, err error) {
	return disp.InvokeWithOptionalArgs(name, com类.DISPATCH_PROPERTYPUTREF, 参数s)
}

// MustPutPropertyRef 变异属性引用或恐慌。
func MustPutPropertyRef(disp *com类.IDispatch, name string, 参数s ...interface{}) (result *com类.I变体结构) {
	r, err := PutPropertyRef(disp, name, 参数s...)
	if err != nil {
		panic(err.Error())
	}
	return r
}

func ForEach(disp *com类.IDispatch, f func(v *com类.I变体结构) error) error {
	newEnum, err := disp.I取属性("_NewEnum")
	if err != nil {
		return err
	}
	defer newEnum.I释放()

	enum, err := newEnum.I取IUnknown对象().IEnumVARIANT(com类.IID_IEnumVariant)
	if err != nil {
		return err
	}
	defer enum.I释放()

	for item, length, err := enum.I下一个(1); length > 0; item, length, err = enum.I下一个(1) {
		if err != nil {
			return err
		}
		if ferr := f(&item); ferr != nil {
			return ferr
		}
	}
	return nil
}
