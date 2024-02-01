//go:build windows
// +build windows

package com子类

import (
	"fmt"
	"strings"
	"testing"
)

// This tests more than one function. It tests all of the functions needed in order to retrieve an
// SafeArray populated with Strings.
func TestSafeArrayConversionString(t *testing.T) {
	I初始化COM库(0)
	defer I取消COM库()

	guid对象, err := I创建GUID对象_从程序ID("QBXMLRP2.RequestProcessor.1")
	if err != nil {
		if err.(*COM错误结构).I取错误码() == CO_E_CLASSSTRING {
			return
		}
		t.Log(err)
		t.FailNow()
	}

	unknown, err := I创建COM对象(guid对象, IID_IUnknown)
	if err != nil {
		t.Log(err)
		t.FailNow()
	}
	defer unknown.I释放()

	dispatch, err := unknown.I查找接口(IID_IDispatch)
	if err != nil {
		t.Log(err)
		t.FailNow()
	}

	var dispid []int32
	dispid, err = dispatch.GetIDsOfName([]string{"OpenConnection2"})
	if err != nil {
		t.Log(err)
		t.FailNow()
	}

	var result *I变体结构
	_, err = dispatch.Invoke(dispid[0], DISPATCH_METHOD, "", "Test Application 1", 1)
	if err != nil {
		t.Log(err)
		t.FailNow()
	}

	dispid, err = dispatch.GetIDsOfName([]string{"BeginSession"})
	if err != nil {
		t.Log(err)
		t.FailNow()
	}

	result, err = dispatch.Invoke(dispid[0], DISPATCH_METHOD, "", 2)
	if err != nil {
		t.Log(err)
		t.FailNow()
	}

	ticket := result.I取文本()

	dispid, err = dispatch.GetIDsOfName([]string{"QBXMLVersionsForSession"})
	if err != nil {
		t.Log(err)
		t.FailNow()
	}

	result, err = dispatch.Invoke(dispid[0], DISPATCH_PROPERTYGET, ticket)
	if err != nil {
		t.Log(err)
		t.FailNow()
	}

	// Where the real tests begin.
	conversion := result.I取数组对象()

	totalElements, _ := conversion.TotalElements(0)
	if totalElements != 13 {
		t.Log(fmt.Sprintf("%d total elements does not equal 13\n", totalElements))
		t.Fail()
	}

	versions := conversion.I取文本数组()
	if len(versions) != 13 {
		t.Log(fmt.Sprintf("%s\n", strings.Join(versions, ", ")))
		t.Fail()
	}

	conversion.I释放()

	dispid, err = dispatch.GetIDsOfName([]string{"EndSession"})
	if err != nil {
		t.Log(err)
		t.FailNow()
	}

	_, err = dispatch.Invoke(dispid[0], DISPATCH_METHOD, ticket)
	if err != nil {
		t.Log(err)
		t.FailNow()
	}

	dispid, err = dispatch.GetIDsOfName([]string{"CloseConnection"})
	if err != nil {
		t.Log(err)
		t.FailNow()
	}

	_, err = dispatch.Invoke(dispid[0], DISPATCH_METHOD)
	if err != nil {
		t.Log(err)
		t.FailNow()
	}
}
