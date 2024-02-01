//go:build windows
// +build windows

package com子类

import "testing"

func TestIEnumVariant_wmi(t *testing.T) {
	var err error
	var classID *GUID

	IID_ISWbemLocator := &GUID{0x76a6415b, 0xcb41, 0x11d1, [8]byte{0x8b, 0x02, 0x00, 0x60, 0x08, 0x06, 0xd9, 0xb6}}

	err = I初始化COM库(0)
	if err != nil {
		t.Errorf("Initialize error: %v", err)
	}
	defer I取消COM库()

	classID, err = I创建GUID对象("WbemScripting.SWbemLocator")
	if err != nil {
		t.Errorf("CreateObject WbemScripting.SWbemLocator returned with %v", err)
	}

	comserver, err := I创建COM对象(classID, IID_IUnknown)
	if err != nil {
		t.Errorf("I创建COM对象 WbemScripting.SWbemLocator returned with %v", err)
	}
	if comserver == nil {
		t.Error("CreateObject WbemScripting.SWbemLocator not an object")
	}
	defer comserver.I释放()

	dispatch, err := comserver.I查找接口(IID_ISWbemLocator)
	if err != nil {
		t.Errorf("context.iunknown.I查找接口 returned with %v", err)
	}
	defer dispatch.I释放()

	wbemServices, err := dispatch.I调用方法("ConnectServer")
	if err != nil {
		t.Errorf("ConnectServer failed with %v", err)
	}
	defer wbemServices.I释放()

	objectset, err := wbemServices.ToIDispatch().I调用方法("ExecQuery", "SELECT * FROM WIN32_Process")
	if err != nil {
		t.Errorf("ExecQuery failed with %v", err)
	}
	defer objectset.I释放()

	enum_property, err := objectset.ToIDispatch().I取属性("_NewEnum")
	if err != nil {
		t.Errorf("Get _NewEnum property failed with %v", err)
	}
	defer enum_property.I释放()

	enum, err := enum_property.I取IUnknown对象().IEnumVARIANT(IID_IEnumVariant)
	if err != nil {
		t.Errorf("IEnumVARIANT() returned with %v", err)
	}
	if enum == nil {
		t.Error("Enum is nil")
		t.FailNow()
	}
	defer enum.I释放()

	for tmp, length, err := enum.I下一个(1); length > 0; tmp, length, err = enum.I下一个(1) {
		if err != nil {
			t.Errorf("I下一个() returned with %v", err)
		}
		tmp_dispatch := tmp.ToIDispatch()
		defer tmp_dispatch.I释放()

		props, err := tmp_dispatch.I取属性("Properties_")
		if err != nil {
			t.Errorf("Get Properties_ property failed with %v", err)
		}
		defer props.I释放()

		props_enum_property, err := props.ToIDispatch().I取属性("_NewEnum")
		if err != nil {
			t.Errorf("Get _NewEnum property failed with %v", err)
		}
		defer props_enum_property.I释放()

		props_enum, err := props_enum_property.I取IUnknown对象().IEnumVARIANT(IID_IEnumVariant)
		if err != nil {
			t.Errorf("IEnumVARIANT failed with %v", err)
		}
		defer props_enum.I释放()

		class_variant, err := tmp_dispatch.I取属性("Name")
		if err != nil {
			t.Errorf("Get Name property failed with %v", err)
		}
		defer class_variant.I释放()

		class_name := class_variant.I取文本()
		t.Logf("Got %v", class_name)
	}
}
