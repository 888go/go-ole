//go:build windows
// +build windows

package com子类

import "testing"

func TestIUnknown(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Error(r)
		}
	}()

	var err error

	err = I初始化COM库(0)
	if err != nil {
		t.Fatal(err)
	}

	defer I取消COM库()

	var unknown *IUnknown

	// oleutil.CreateObject()
	unknown, err = I创建COM对象(CLSID_COMEchoTestObject, IID_IUnknown)
	if err != nil {
		t.Fatal(err)
		return
	}
	unknown.I释放()
}
