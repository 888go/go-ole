package com子类

import (
	"unicode/utf16"
	"unsafe"
)

// I创建GUID对象 ClassIDFrom，无论给定的是程序ID还是应用程序字符串。
//
// 帮助程序，用于检查程序ID中的类ID和字符串中的类标识. 如果您知道您正在使用哪个功能，使用单独的功能会更快，但这将检查您的可用功能。
func I创建GUID对象(programID string) (guid对象 *GUID, err error) {
	guid对象, err = I创建GUID对象_从程序ID(programID)
	if err != nil {
		guid对象, err = I创建GUID对象_从标识(programID)
		if err != nil {
			return
		}
	}
	return
}

// I字节指针到文本 BytePtrToString 将字节指针转换为Go字符串。
func I字节指针到文本(p *byte) string {
	a := (*[10000]uint8)(unsafe.Pointer(p))
	i := 0
	for a[i] != 0 {
		i++
	}
	return string(a[:i])
}

// UTF16PtrToString别名  UTF16PtrToString 是别名LpOleStrToString。
//
// 出于兼容性原因而保留。
func UTF16PtrToString别名(p *uint16) string {
	return Unicode指针到文本(p)
}

// Unicode指针到文本  LpOleStrToString 将COM Unicode转换为Go字符串。
func Unicode指针到文本(p *uint16) string {
	if p == nil {
		return ""
	}

	length := lpOleStrLen(p)
	a := make([]uint16, length)

	ptr := unsafe.Pointer(p)

	for i := 0; i < int(length); i++ {
		a[i] = *(*uint16)(ptr)
		ptr = unsafe.Pointer(uintptr(ptr) + 2)
	}

	return string(utf16.Decode(a))
}

// BstrToString 将COM二进制字符串转换为Go字符串。
func BstrToString(p *uint16) string {
	if p == nil {
		return ""
	}
	length := SysStringLen((*int16)(unsafe.Pointer(p)))
	a := make([]uint16, length)

	ptr := unsafe.Pointer(p)

	for i := 0; i < int(length); i++ {
		a[i] = *(*uint16)(ptr)
		ptr = unsafe.Pointer(uintptr(ptr) + 2)
	}
	return string(utf16.Decode(a))
}

// lpOleStrLen 返回Unicode字符串的长度。
func lpOleStrLen(p *uint16) (length int64) {
	if p == nil {
		return 0
	}

	ptr := unsafe.Pointer(p)

	for i := 0; ; i++ {
		if 0 == *(*uint16)(ptr) {
			length = int64(i)
			break
		}
		ptr = unsafe.Pointer(uintptr(ptr) + 2)
	}
	return
}

// convertHresultToError 如果调用不成功，则将syscall转换为error。
func convertHresultToError(hr uintptr, r2 uintptr, ignore error) (err error) {
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}
