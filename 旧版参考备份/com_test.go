//go:build windows
// +build windows

package com子类

import (
	"fmt"
	"testing"
)

func TestComSetupAndShutDown(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Log(r)
			t.Fail()
		}
	}()

	err := coInitialize()
	if err != nil {
		t.Error(err)
		t.FailNow()
	}

	I取消COM库()
}

func TestComPublicSetupAndShutDown(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Log(r)
			t.Fail()
		}
	}()

	err := I初始化COM库(0)
	if err != nil {
		t.Error(err)
		t.FailNow()
	}

	I取消COM库()
}

func TestComPublicSetupAndShutDown_WithValue(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Log(r)
			t.Fail()
		}
	}()

	err := I初始化COM库(5)
	if err != nil {
		t.Error(err)
		t.FailNow()
	}

	I取消COM库()
}

func TestComExSetupAndShutDown(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Log(r)
			t.Fail()
		}
	}()

	err := coInitializeEx(COINIT_MULTITHREADED)
	if err != nil {
		t.Error(err)
		t.FailNow()
	}

	I取消COM库()
}

func TestComPublicExSetupAndShutDown(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Log(r)
			t.Fail()
		}
	}()

	err := I初始化COM库_线程(0, COINIT_MULTITHREADED)
	if err != nil {
		t.Error(err)
		t.FailNow()
	}

	I取消COM库()
}

func TestComPublicExSetupAndShutDown_WithValue(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Log(r)
			t.Fail()
		}
	}()

	err := I初始化COM库_线程(5, COINIT_MULTITHREADED)
	if err != nil {
		t.Error(err)
		t.FailNow()
	}

	I取消COM库()
}

func TestClsidFromProgID_WindowsMediaNSSManager(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Log(r)
			t.Fail()
		}
	}()

	expected := &GUID{0x92498132, 0x4D1A, 0x4297, [8]byte{0x9B, 0x78, 0x9E, 0x2E, 0x4B, 0xA9, 0x9C, 0x07}}

	coInitialize()
	defer I取消COM库()
	actual, err := I创建GUID对象_从程序ID("WMPNSSCI.NSSManager")
	if err == nil {
		if !I是否相等_GUID(expected, actual) {
			t.Log(err)
			t.Log(fmt.Sprintf("Actual GUID: %+v\n", actual))
			t.Fail()
		}
	}
}

func TestClsidFromString_WindowsMediaNSSManager(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Log(r)
			t.Fail()
		}
	}()

	expected := &GUID{0x92498132, 0x4D1A, 0x4297, [8]byte{0x9B, 0x78, 0x9E, 0x2E, 0x4B, 0xA9, 0x9C, 0x07}}

	coInitialize()
	defer I取消COM库()
	actual, err := I创建GUID对象_从标识("{92498132-4D1A-4297-9B78-9E2E4BA99C07}")

	if !I是否相等_GUID(expected, actual) {
		t.Log(err)
		t.Log(fmt.Sprintf("Actual GUID: %+v\n", actual))
		t.Fail()
	}
}

func TestCreateInstance_WindowsMediaNSSManager(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Log(r)
			t.Fail()
		}
	}()

	expected := &GUID{0x92498132, 0x4D1A, 0x4297, [8]byte{0x9B, 0x78, 0x9E, 0x2E, 0x4B, 0xA9, 0x9C, 0x07}}

	coInitialize()
	defer I取消COM库()
	actual, err := I创建GUID对象_从程序ID("WMPNSSCI.NSSManager")

	if err == nil {
		if !I是否相等_GUID(expected, actual) {
			t.Log(err)
			t.Log(fmt.Sprintf("Actual GUID: %+v\n", actual))
			t.Fail()
		}

		unknown, err := I创建COM对象(actual, IID_IUnknown)
		if err != nil {
			t.Log(err)
			t.Fail()
		}
		unknown.I释放()
	}
}

func TestError(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			t.Log(r)
			t.Fail()
		}
	}()

	coInitialize()
	defer I取消COM库()
	_, err := I创建GUID对象_从程序ID("INTERFACE-NOT-FOUND")
	if err == nil {
		t.Fatalf("should be fail: %v", err)
	}

	switch vt := err.(type) {
	case *COM错误结构:
	default:
		t.Fatalf("should be *ole.COM错误结构 %t", vt)
	}
}
