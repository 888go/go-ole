//go:build windows
// +build windows

package com子类

import (
	"unsafe"
)

var (
	procSafeArrayAccessData        = modoleaut32.NewProc("SafeArrayAccessData")
	procSafeArrayAllocData         = modoleaut32.NewProc("SafeArrayAllocData")
	procSafeArrayAllocDescriptor   = modoleaut32.NewProc("SafeArrayAllocDescriptor")
	procSafeArrayAllocDescriptorEx = modoleaut32.NewProc("SafeArrayAllocDescriptorEx")
	procSafeArrayCopy              = modoleaut32.NewProc("SafeArrayCopy")
	procSafeArrayCopyData          = modoleaut32.NewProc("SafeArrayCopyData")
	procSafeArrayCreate            = modoleaut32.NewProc("SafeArrayCreate")
	procSafeArrayCreateEx          = modoleaut32.NewProc("SafeArrayCreateEx")
	procSafeArrayCreateVector      = modoleaut32.NewProc("SafeArrayCreateVector")
	procSafeArrayCreateVectorEx    = modoleaut32.NewProc("SafeArrayCreateVectorEx")
	procSafeArrayDestroy           = modoleaut32.NewProc("SafeArrayDestroy")
	procSafeArrayDestroyData       = modoleaut32.NewProc("SafeArrayDestroyData")
	procSafeArrayDestroyDescriptor = modoleaut32.NewProc("SafeArrayDestroyDescriptor")
	procSafeArrayGetDim            = modoleaut32.NewProc("SafeArrayGetDim")
	procSafeArrayGetElement        = modoleaut32.NewProc("SafeArrayGetElement")
	procSafeArrayGetElemsize       = modoleaut32.NewProc("SafeArrayGetElemsize")
	procSafeArrayGetIID            = modoleaut32.NewProc("SafeArrayGetIID")
	procSafeArrayGetLBound         = modoleaut32.NewProc("SafeArrayGetLBound")
	procSafeArrayGetUBound         = modoleaut32.NewProc("SafeArrayGetUBound")
	procSafeArrayGetVartype        = modoleaut32.NewProc("SafeArrayGetVartype")
	procSafeArrayLock              = modoleaut32.NewProc("SafeArrayLock")
	procSafeArrayPtrOfIndex        = modoleaut32.NewProc("SafeArrayPtrOfIndex")
	procSafeArrayUnaccessData      = modoleaut32.NewProc("SafeArrayUnaccessData")
	procSafeArrayUnlock            = modoleaut32.NewProc("SafeArrayUnlock")
	procSafeArrayPutElement        = modoleaut32.NewProc("SafeArrayPutElement")
	//procSafeArrayRedim             = modoleaut32.NewProc("SafeArrayRedim") // TODO
	//procSafeArraySetIID            = modoleaut32.NewProc("SafeArraySetIID") // TODO
	procSafeArrayGetRecordInfo = modoleaut32.NewProc("SafeArrayGetRecordInfo")
	procSafeArraySetRecordInfo = modoleaut32.NewProc("SafeArraySetRecordInfo")
)

// safeArrayAccessData 返回原始数组指针。
//
// AKA: Windows API中的SafeArrayAccessData。
func safeArrayAccessData(safearray *SafeArray) (element uintptr, err error) {
	err = convertHresultToError(
		procSafeArrayAccessData.Call(
			uintptr(unsafe.Pointer(safearray)),
			uintptr(unsafe.Pointer(&element))))
	return
}

// safeArrayUnaccessData 释放原始阵列。
//
// AKA：Windows API中的SafeArrayUnaccessData。
func safeArrayUnaccessData(safearray *SafeArray) (err error) {
	err = convertHresultToError(procSafeArrayUnaccessData.Call(uintptr(unsafe.Pointer(safearray))))
	return
}

// safeArrayAllocData 分配SafeArray。
//
// AKA:Windows API中的SafeArrayAllocData。
func safeArrayAllocData(safearray *SafeArray) (err error) {
	err = convertHresultToError(procSafeArrayAllocData.Call(uintptr(unsafe.Pointer(safearray))))
	return
}

// safeArrayAllocDescriptor 分配SafeArray。
//
// AKA:Windows API中的SafeArrayAllocDescriptor。
func safeArrayAllocDescriptor(dimensions uint32) (safearray *SafeArray, err error) {
	err = convertHresultToError(
		procSafeArrayAllocDescriptor.Call(uintptr(dimensions), uintptr(unsafe.Pointer(&safearray))))
	return
}

// safeArrayAllocDescriptorEx 分配SafeArray。
//
// AKA:Windows API中的SafeArrayAllocDescriptorEx。
func safeArrayAllocDescriptorEx(variantType VT, dimensions uint32) (safearray *SafeArray, err error) {
	err = convertHresultToError(
		procSafeArrayAllocDescriptorEx.Call(
			uintptr(variantType),
			uintptr(dimensions),
			uintptr(unsafe.Pointer(&safearray))))
	return
}

// safeArrayCopy 返回SafeArray的副本。
//
// AKA：Windows API中的SafeArrayCopy。
func safeArrayCopy(original *SafeArray) (safearray *SafeArray, err error) {
	err = convertHresultToError(
		procSafeArrayCopy.Call(
			uintptr(unsafe.Pointer(original)),
			uintptr(unsafe.Pointer(&safearray))))
	return
}

// safeArrayCopyData 将SafeArray复制到另一个SafeArray对象中。
//
// AKA:Windows API中的SafeArrayCopyData。
func safeArrayCopyData(original *SafeArray, duplicate *SafeArray) (err error) {
	err = convertHresultToError(
		procSafeArrayCopyData.Call(
			uintptr(unsafe.Pointer(original)),
			uintptr(unsafe.Pointer(duplicate))))
	return
}

// safeArrayCreate 创建SafeArray。
//
// AKA:Windows API中的SafeArrayCreate。
func safeArrayCreate(variantType VT, dimensions uint32, bounds *SafeArrayBound) (safearray *SafeArray, err error) {
	sa, _, err := procSafeArrayCreate.Call(
		uintptr(variantType),
		uintptr(dimensions),
		uintptr(unsafe.Pointer(bounds)))
	safearray = (*SafeArray)(unsafe.Pointer(&sa))
	return
}

// safeArrayCreateEx 创建SafeArray。
//
// AKA:Windows API中的SafeArrayCreateEx。
func safeArrayCreateEx(variantType VT, dimensions uint32, bounds *SafeArrayBound, extra uintptr) (safearray *SafeArray, err error) {
	sa, _, err := procSafeArrayCreateEx.Call(
		uintptr(variantType),
		uintptr(dimensions),
		uintptr(unsafe.Pointer(bounds)),
		extra)
	safearray = (*SafeArray)(unsafe.Pointer(sa))
	return
}

// safeArrayCreateVector 创建SafeArray。
//
// AKA:Windows API中的SafeArrayCreateVector。
func safeArrayCreateVector(variantType VT, lowerBound int32, length uint32) (safearray *SafeArray, err error) {
	sa, _, err := procSafeArrayCreateVector.Call(
		uintptr(variantType),
		uintptr(lowerBound),
		uintptr(length))
	safearray = (*SafeArray)(unsafe.Pointer(sa))
	return
}

// safeArrayCreateVectorEx 创建SafeArray。
//
// AKA:Windows API中的SafeArrayCreateVectorEx。
func safeArrayCreateVectorEx(variantType VT, lowerBound int32, length uint32, extra uintptr) (safearray *SafeArray, err error) {
	sa, _, err := procSafeArrayCreateVectorEx.Call(
		uintptr(variantType),
		uintptr(lowerBound),
		uintptr(length),
		extra)
	safearray = (*SafeArray)(unsafe.Pointer(sa))
	return
}

// safeArrayDestroy 销毁SafeArray对象。
//
// AKA：Windows API中的SafeArrayDestroy。
func safeArrayDestroy(safearray *SafeArray) (err error) {
	err = convertHresultToError(procSafeArrayDestroy.Call(uintptr(unsafe.Pointer(safearray))))
	return
}

// safeArrayDestroyData 销毁SafeArray对象。
//
// AKA:Windows API中的SafeArrayDestroyData。
func safeArrayDestroyData(safearray *SafeArray) (err error) {
	err = convertHresultToError(procSafeArrayDestroyData.Call(uintptr(unsafe.Pointer(safearray))))
	return
}

// safeArrayDestroyDescriptor 销毁SafeArray对象。
//
// AKA:Windows API中的SafeArrayDestroyDescriptor。
func safeArrayDestroyDescriptor(safearray *SafeArray) (err error) {
	err = convertHresultToError(procSafeArrayDestroyDescriptor.Call(uintptr(unsafe.Pointer(safearray))))
	return
}

// safeArrayGetDim 是SafeArray中的维数。
//
// SafeArrays可以有多个维度。也就是说，它可以是多维数组。
//
// AKA:Windows API中的SafeArrayGetDim。
func safeArrayGetDim(safearray *SafeArray) (dimensions *uint32, err error) {
	l, _, err := procSafeArrayGetDim.Call(uintptr(unsafe.Pointer(safearray)))
	dimensions = (*uint32)(unsafe.Pointer(l))
	return
}

// safeArrayGetElementSize 是以字节为单位的元素大小。
//
// AKA:Windows API中的SafeArrayGetElemsize。
func safeArrayGetElementSize(safearray *SafeArray) (length *uint32, err error) {
	l, _, err := procSafeArrayGetElemsize.Call(uintptr(unsafe.Pointer(safearray)))
	length = (*uint32)(unsafe.Pointer(l))
	return
}

// safeArrayGetElement 检索给定索引处的元素。
func safeArrayGetElement(safearray *SafeArray, index int32, pv unsafe.Pointer) error {
	return convertHresultToError(
		procSafeArrayGetElement.Call(
			uintptr(unsafe.Pointer(safearray)),
			uintptr(unsafe.Pointer(&index)),
			uintptr(pv)))
}

// safeArrayGetElementString 检索给定索引处的元素并转换为字符串。
func safeArrayGetElementString(safearray *SafeArray, index int32) (str string, err error) {
	var element *int16
	err = convertHresultToError(
		procSafeArrayGetElement.Call(
			uintptr(unsafe.Pointer(safearray)),
			uintptr(unsafe.Pointer(&index)),
			uintptr(unsafe.Pointer(&element))))
	str = BstrToString(*(**uint16)(unsafe.Pointer(&element)))
	SysFreeString(element)
	return
}

// safeArrayGetIID 是SafeArray中元素的InterfaceID。
//
// AKA:Windows API中的SafeArrayGetIID。
func safeArrayGetIID(safearray *SafeArray) (guid *GUID, err error) {
	err = convertHresultToError(
		procSafeArrayGetIID.Call(
			uintptr(unsafe.Pointer(safearray)),
			uintptr(unsafe.Pointer(&guid))))
	return
}

// safeArrayGetLBound 返回SafeArray的下限。
//
// SafeArrays可以有多个维度。也就是说，它可以是多维数组。
//
// AKA:Windows API中的SafeArrayGetLBound。
func safeArrayGetLBound(safearray *SafeArray, dimension uint32) (lowerBound int32, err error) {
	err = convertHresultToError(
		procSafeArrayGetLBound.Call(
			uintptr(unsafe.Pointer(safearray)),
			uintptr(dimension),
			uintptr(unsafe.Pointer(&lowerBound))))
	return
}

// safeArrayGetUBound 返回SafeArray的上限。
//
// SafeArrays可以有多个维度。也就是说，它可以是多维数组。
//
// AKA:Windows API中的SafeArrayGetUBound。
func safeArrayGetUBound(safearray *SafeArray, dimension uint32) (upperBound int32, err error) {
	err = convertHresultToError(
		procSafeArrayGetUBound.Call(
			uintptr(unsafe.Pointer(safearray)),
			uintptr(dimension),
			uintptr(unsafe.Pointer(&upperBound))))
	return
}

// safeArrayGetVartype 返回SafeArray的数据类型。
//
// AKA:Windows API中的SafeArrayGetVartype。
func safeArrayGetVartype(safearray *SafeArray) (varType uint16, err error) {
	err = convertHresultToError(
		procSafeArrayGetVartype.Call(
			uintptr(unsafe.Pointer(safearray)),
			uintptr(unsafe.Pointer(&varType))))
	return
}

// safeArrayLock 锁定SafeArray进行读取以修改SafeArray。
//
// 这必须在某些调用期间调用，以确保另一个进程在编辑期间不会读取或写入SafeArray。
//
// AKA：Windows API中的SafeArrayLock。
func safeArrayLock(safearray *SafeArray) (err error) {
	err = convertHresultToError(procSafeArrayLock.Call(uintptr(unsafe.Pointer(safearray))))
	return
}

// safeArrayUnlock 解锁SafeArray进行读取。
//
// AKA:Windows API中的SafeArrayUnlock。
func safeArrayUnlock(safearray *SafeArray) (err error) {
	err = convertHresultToError(procSafeArrayUnlock.Call(uintptr(unsafe.Pointer(safearray))))
	return
}

// safeArrayPutElement 将数据元素存储在数组中的指定位置。
//
// AKA:Windows API中的SafeArrayPutElement。
func safeArrayPutElement(safearray *SafeArray, index int64, element uintptr) (err error) {
	err = convertHresultToError(
		procSafeArrayPutElement.Call(
			uintptr(unsafe.Pointer(safearray)),
			uintptr(unsafe.Pointer(&index)),
			uintptr(unsafe.Pointer(element))))
	return
}

// safeArrayGetRecordInfo 自定义类型的IRecordInfo信息访问。
//
// AKA:Windows API中的SafeArrayGetRecordInfo。
//
// XXX: 必须实现IRecordInfo接口才能返回。
func safeArrayGetRecordInfo(safearray *SafeArray) (recordInfo interface{}, err error) {
	err = convertHresultToError(
		procSafeArrayGetRecordInfo.Call(
			uintptr(unsafe.Pointer(safearray)),
			uintptr(unsafe.Pointer(&recordInfo))))
	return
}

// safeArraySetRecordInfo 更改自定义类型的IRecordInfo信息。
//
// AKA:Windows API中的SafeArraySetRecordInfo。
//
// XXX:必须实现IRecordInfo接口才能返回。
func safeArraySetRecordInfo(safearray *SafeArray, recordInfo interface{}) (err error) {
	err = convertHresultToError(
		procSafeArraySetRecordInfo.Call(
			uintptr(unsafe.Pointer(safearray)),
			uintptr(unsafe.Pointer(&recordInfo))))
	return
}
