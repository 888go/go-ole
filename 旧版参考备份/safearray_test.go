package com子类

// 这将测试多个功能。它测试检索用字符串填充的SafeArray所需的所有函数。
func Example_safeArrayGetElementString() {
	I初始化COM库(0)
	defer I取消COM库()

	guid对象, err := I创建GUID对象_从程序ID("QBXMLRP2.RequestProcessor.1")
	if err != nil {
		if err.(*COM错误结构).I取错误码() == CO_E_CLASSSTRING {
			return
		}
	}

	unknown, err := I创建COM对象(guid对象, IID_IUnknown)
	if err != nil {
		return
	}
	defer unknown.I释放()

	dispatch, err := unknown.I查找接口(IID_IDispatch)
	if err != nil {
		return
	}

	var dispid []int32
	dispid, err = dispatch.GetIDsOfName([]string{"OpenConnection2"})
	if err != nil {
		return
	}

	var result *I变体结构
	_, err = dispatch.Invoke(dispid[0], DISPATCH_METHOD, "", "Test Application 1", 1)
	if err != nil {
		return
	}

	dispid, err = dispatch.GetIDsOfName([]string{"BeginSession"})
	if err != nil {
		return
	}

	result, err = dispatch.Invoke(dispid[0], DISPATCH_METHOD, "", 2)
	if err != nil {
		return
	}

	ticket := result.I取文本()

	dispid, err = dispatch.GetIDsOfName([]string{"QBXMLVersionsForSession"})
	if err != nil {
		return
	}

	result, err = dispatch.Invoke(dispid[0], DISPATCH_PROPERTYGET, ticket)
	if err != nil {
		return
	}

	// Where the real tests begin.
	var qbXMLVersions *SafeArray
	var qbXmlVersionStrings []string
	qbXMLVersions = result.I取数组对象().Array

	// Get array bounds
	var LowerBounds int32
	var UpperBounds int32
	LowerBounds, err = safeArrayGetLBound(qbXMLVersions, 1)
	if err != nil {
		return
	}

	UpperBounds, err = safeArrayGetUBound(qbXMLVersions, 1)
	if err != nil {
		return
	}

	totalElements := UpperBounds - LowerBounds + 1
	qbXmlVersionStrings = make([]string, totalElements)

	for i := int32(0); i < totalElements; i++ {
		qbXmlVersionStrings[i], _ = safeArrayGetElementString(qbXMLVersions, i)
	}

	// 释放安全阵列内存
	safeArrayDestroy(qbXMLVersions)

	dispid, err = dispatch.GetIDsOfName([]string{"EndSession"})
	if err != nil {
		return
	}

	_, err = dispatch.Invoke(dispid[0], DISPATCH_METHOD, ticket)
	if err != nil {
		return
	}

	dispid, err = dispatch.GetIDsOfName([]string{"CloseConnection"})
	if err != nil {
		return
	}

	_, err = dispatch.Invoke(dispid[0], DISPATCH_METHOD)
	if err != nil {
		return
	}
}
