//go:build windows
// +build windows

package com子类

import "testing"

func wrapCOMExecute(t *testing.T, callback func(*testing.T)) {
	defer func() {
		if r := recover(); r != nil {
			t.Error(r)
		}
	}()

	err := I初始化COM库(0)
	if err != nil {
		t.Fatal(err)
	}
	defer I取消COM库()

	callback(t)
}

func wrapDispatch(t *testing.T, ClassID, UnknownInterfaceID, DispatchInterfaceID *GUID, callback func(*testing.T, *IUnknown, *IDispatch)) {
	var unknown *IUnknown
	var dispatch *IDispatch
	var err error

	unknown, err = I创建COM对象(ClassID, UnknownInterfaceID)
	if err != nil {
		t.Error(err)
		return
	}
	defer unknown.I释放()

	dispatch, err = unknown.I查找接口(DispatchInterfaceID)
	if err != nil {
		t.Error(err)
		return
	}
	defer dispatch.I释放()

	callback(t, unknown, dispatch)
}

func wrapGoOLETestCOMServerEcho(t *testing.T, callback func(*testing.T, *IUnknown, *IDispatch)) {
	wrapCOMExecute(t, func(t *testing.T) {
		wrapDispatch(t, CLSID_COMEchoTestObject, IID_IUnknown, IID_ICOMEchoTestObject, callback)
	})
}

func wrapGoOLETestCOMServerScalar(t *testing.T, callback func(*testing.T, *IUnknown, *IDispatch)) {
	wrapCOMExecute(t, func(t *testing.T) {
		wrapDispatch(t, CLSID_COMTestScalarClass, IID_IUnknown, IID_ICOMTestTypes, callback)
	})
}

func TestIDispatch_goolecomserver_stringfield(t *testing.T) {
	wrapGoOLETestCOMServerScalar(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "StringField"
		expected := "Test I取文本"
		_, err := idispatch.I设置属性(method, expected)
		if err != nil {
			t.Error(err)
			return
		}
		variant, err := idispatch.I取属性(method)
		if err != nil {
			t.Error(err)
			return
		}
		defer variant.I释放()
		actual, passed := variant.I取interface().(string)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "string", variant.VT, variant.Val)
			return
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_int8field(t *testing.T) {
	wrapGoOLETestCOMServerScalar(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "Int8Field"
		expected := int8(2)
		_, err := idispatch.I设置属性(method, expected)
		if err != nil {
			t.Error(err)
			return
		}
		variant, err := idispatch.I取属性(method)
		if err != nil {
			t.Error(err)
			return
		}
		defer variant.I释放()
		actual, passed := variant.I取interface().(int8)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "int8", variant.VT, variant.Val)
			return
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_uint8field(t *testing.T) {
	wrapGoOLETestCOMServerScalar(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "UInt8Field"
		expected := uint8(4)
		_, err := idispatch.I设置属性(method, expected)
		if err != nil {
			t.Error(err)
			return
		}
		variant, err := idispatch.I取属性(method)
		if err != nil {
			t.Error(err)
			return
		}
		defer variant.I释放()
		actual, passed := variant.I取interface().(uint8)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "uint8", variant.VT, variant.Val)
			return
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_int16field(t *testing.T) {
	wrapGoOLETestCOMServerScalar(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "Int16Field"
		expected := int16(4)
		_, err := idispatch.I设置属性(method, expected)
		if err != nil {
			t.Error(err)
			return
		}
		variant, err := idispatch.I取属性(method)
		if err != nil {
			t.Error(err)
			return
		}
		defer variant.I释放()
		actual, passed := variant.I取interface().(int16)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "int16", variant.VT, variant.Val)
			return
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_uint16field(t *testing.T) {
	wrapGoOLETestCOMServerScalar(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "UInt16Field"
		expected := uint16(4)
		_, err := idispatch.I设置属性(method, expected)
		if err != nil {
			t.Error(err)
			return
		}
		variant, err := idispatch.I取属性(method)
		if err != nil {
			t.Error(err)
			return
		}
		defer variant.I释放()
		actual, passed := variant.I取interface().(uint16)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "uint16", variant.VT, variant.Val)
			return
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_int32field(t *testing.T) {
	wrapGoOLETestCOMServerScalar(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "Int32Field"
		expected := int32(8)
		_, err := idispatch.I设置属性(method, expected)
		if err != nil {
			t.Error(err)
			return
		}
		variant, err := idispatch.I取属性(method)
		if err != nil {
			t.Error(err)
			return
		}
		defer variant.I释放()
		actual, passed := variant.I取interface().(int32)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "int32", variant.VT, variant.Val)
			return
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_uint32field(t *testing.T) {
	wrapGoOLETestCOMServerScalar(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "UInt32Field"
		expected := uint32(16)
		_, err := idispatch.I设置属性(method, expected)
		if err != nil {
			t.Error(err)
			return
		}
		variant, err := idispatch.I取属性(method)
		if err != nil {
			t.Error(err)
			return
		}
		defer variant.I释放()
		actual, passed := variant.I取interface().(uint32)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "uint32", variant.VT, variant.Val)
			return
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_int64field(t *testing.T) {
	wrapGoOLETestCOMServerScalar(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "Int64Field"
		expected := int64(32)
		_, err := idispatch.I设置属性(method, expected)
		if err != nil {
			t.Error(err)
			return
		}
		variant, err := idispatch.I取属性(method)
		if err != nil {
			t.Error(err)
			return
		}
		defer variant.I释放()
		actual, passed := variant.I取interface().(int64)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "int64", variant.VT, variant.Val)
			return
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_uint64field(t *testing.T) {
	wrapGoOLETestCOMServerScalar(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "UInt64Field"
		expected := uint64(64)
		_, err := idispatch.I设置属性(method, expected)
		if err != nil {
			t.Error(err)
			return
		}
		variant, err := idispatch.I取属性(method)
		if err != nil {
			t.Error(err)
			return
		}
		defer variant.I释放()
		actual, passed := variant.I取interface().(uint64)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "uint64", variant.VT, variant.Val)
			return
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_booleanfield_true(t *testing.T) {
	wrapGoOLETestCOMServerScalar(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "BooleanField"
		expected := true
		_, err := idispatch.I设置属性(method, expected)
		if err != nil {
			t.Error(err)
			return
		}
		variant, err := idispatch.I取属性(method)
		if err != nil {
			t.Error(err)
			return
		}
		defer variant.I释放()
		actual, passed := variant.I取interface().(bool)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "bool", variant.VT, variant.Val)
			return
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_booleanfield_false(t *testing.T) {
	wrapGoOLETestCOMServerScalar(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "BooleanField"
		expected := false
		_, err := idispatch.I设置属性(method, expected)
		if err != nil {
			t.Error(err)
			return
		}
		variant, err := idispatch.I取属性(method)
		if err != nil {
			t.Error(err)
			return
		}
		defer variant.I释放()
		actual, passed := variant.I取interface().(bool)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "bool", variant.VT, variant.Val)
			return
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_float32field(t *testing.T) {
	wrapGoOLETestCOMServerScalar(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "Float32Field"
		expected := float32(2.2)
		_, err := idispatch.I设置属性(method, expected)
		if err != nil {
			t.Error(err)
		}
		variant, err := idispatch.I取属性(method)
		if err != nil {
			t.Error(err)
			return
		}
		defer variant.I释放()
		actual, passed := variant.I取interface().(float32)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "float32", variant.VT, variant.Val)
			return
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_float64field(t *testing.T) {
	wrapGoOLETestCOMServerScalar(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "Float64Field"
		expected := float64(4.4)
		_, err := idispatch.I设置属性(method, expected)
		if err != nil {
			t.Error(err)
		}
		variant, err := idispatch.I取属性(method)
		if err != nil {
			t.Error(err)
			return
		}
		defer variant.I释放()
		actual, passed := variant.I取interface().(float64)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "float64", variant.VT, variant.Val)
			return
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_echostring(t *testing.T) {
	wrapGoOLETestCOMServerEcho(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "EchoString"
		expected := "Test I取文本"
		variant, err := idispatch.I调用方法(method, expected)
		if err != nil {
			t.Error(err)
			return
		}
		defer variant.I释放()
		actual, passed := variant.I取interface().(string)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "string", variant.VT, variant.Val)
			return
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_echoboolean(t *testing.T) {
	wrapGoOLETestCOMServerEcho(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "EchoBoolean"
		expected := true
		variant, err := idispatch.I调用方法(method, expected)
		if err != nil {
			t.Error(err)
			return
		}
		defer variant.I释放()
		actual, passed := variant.I取interface().(bool)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "bool", variant.VT, variant.Val)
			return
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_echoint8(t *testing.T) {
	wrapGoOLETestCOMServerEcho(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "EchoInt8"
		expected := int8(1)
		variant, err := idispatch.I调用方法(method, expected)
		if err != nil {
			t.Error(err)
			return
		}
		defer variant.I释放()
		actual, passed := variant.I取interface().(int8)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "int8", variant.VT, variant.Val)
			return
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_echouint8(t *testing.T) {
	wrapGoOLETestCOMServerEcho(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "EchoUInt8"
		expected := uint8(1)
		variant, err := idispatch.I调用方法(method, expected)
		if err != nil {
			t.Error(err)
			return
		}
		defer variant.I释放()
		actual, passed := variant.I取interface().(uint8)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "uint8", variant.VT, variant.Val)
			return
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_echoint16(t *testing.T) {
	wrapGoOLETestCOMServerEcho(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "EchoInt16"
		expected := int16(1)
		variant, err := idispatch.I调用方法(method, expected)
		if err != nil {
			t.Error(err)
			return
		}
		defer variant.I释放()
		actual, passed := variant.I取interface().(int16)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "int16", variant.VT, variant.Val)
			return
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_echouint16(t *testing.T) {
	wrapGoOLETestCOMServerEcho(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "EchoUInt16"
		expected := uint16(1)
		variant, err := idispatch.I调用方法(method, expected)
		if err != nil {
			t.Error(err)
			return
		}
		defer variant.I释放()
		actual, passed := variant.I取interface().(uint16)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "uint16", variant.VT, variant.Val)
			return
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_echoint32(t *testing.T) {
	wrapGoOLETestCOMServerEcho(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "EchoInt32"
		expected := int32(2)
		variant, err := idispatch.I调用方法(method, expected)
		if err != nil {
			t.Error(err)
			return
		}
		defer variant.I释放()
		actual, passed := variant.I取interface().(int32)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "int32", variant.VT, variant.Val)
			return
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_echouint32(t *testing.T) {
	wrapGoOLETestCOMServerEcho(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "EchoUInt32"
		expected := uint32(4)
		variant, err := idispatch.I调用方法(method, expected)
		if err != nil {
			t.Error(err)
			return
		}
		defer variant.I释放()
		actual, passed := variant.I取interface().(uint32)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "uint32", variant.VT, variant.Val)
			return
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_echoint64(t *testing.T) {
	wrapGoOLETestCOMServerEcho(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "EchoInt64"
		expected := int64(1)
		variant, err := idispatch.I调用方法(method, expected)
		if err != nil {
			t.Error(err)
			return
		}
		defer variant.I释放()
		actual, passed := variant.I取interface().(int64)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "int64", variant.VT, variant.Val)
			return
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_echouint64(t *testing.T) {
	wrapGoOLETestCOMServerEcho(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "EchoUInt64"
		expected := uint64(1)
		variant, err := idispatch.I调用方法(method, expected)
		if err != nil {
			t.Error(err)
			return
		}
		defer variant.I释放()
		actual, passed := variant.I取interface().(uint64)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "uint64", variant.VT, variant.Val)
			return
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_echofloat32(t *testing.T) {
	wrapGoOLETestCOMServerEcho(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "EchoFloat32"
		expected := float32(2.2)
		variant, err := idispatch.I调用方法(method, expected)
		if err != nil {
			t.Error(err)
			return
		}
		defer variant.I释放()
		actual, passed := variant.I取interface().(float32)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "float32", variant.VT, variant.Val)
			return
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}

func TestIDispatch_goolecomserver_echofloat64(t *testing.T) {
	wrapGoOLETestCOMServerEcho(t, func(t *testing.T, unknown *IUnknown, idispatch *IDispatch) {
		method := "EchoFloat64"
		expected := float64(2.2)
		variant, err := idispatch.I调用方法(method, expected)
		if err != nil {
			t.Error(err)
			return
		}
		defer variant.I释放()
		actual, passed := variant.I取interface().(float64)
		if !passed {
			t.Errorf("%s() did not convert to %s, variant is %s with %v value", method, "float64", variant.VT, variant.Val)
			return
		}
		if actual != expected {
			t.Errorf("%s() expected %v did not match %v", method, expected, actual)
		}
	})
}
