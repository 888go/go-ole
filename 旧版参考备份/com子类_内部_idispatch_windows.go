//go:build windows
// +build windows

package com子类

import (
	"math/big"
	"syscall"
	"time"
	"unsafe"
)

func getIDsOfName(disp *IDispatch, names []string) (dispid []int32, err error) {
	wnames := make([]*uint16, len(names))
	for i := 0; i < len(names); i++ {
		wnames[i] = syscall.StringToUTF16Ptr(names[i])
	}
	dispid = make([]int32, len(names))
	namelen := uint32(len(names))
	hr, _, _ := syscall.Syscall6(
		disp.VTable().GetIDsOfNames,
		6,
		uintptr(unsafe.Pointer(disp)),
		uintptr(unsafe.Pointer(IID_NULL)),
		uintptr(unsafe.Pointer(&wnames[0])),
		uintptr(namelen),
		uintptr(I取用户默认LCID()),
		uintptr(unsafe.Pointer(&dispid[0])))
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}

func getTypeInfoCount(disp *IDispatch) (c uint32, err error) {
	hr, _, _ := syscall.Syscall(
		disp.VTable().GetTypeInfoCount,
		2,
		uintptr(unsafe.Pointer(disp)),
		uintptr(unsafe.Pointer(&c)),
		0)
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}

func getTypeInfo(disp *IDispatch) (tinfo *ITypeInfo, err error) {
	hr, _, _ := syscall.Syscall(
		disp.VTable().GetTypeInfo,
		3,
		uintptr(unsafe.Pointer(disp)),
		uintptr(I取用户默认LCID()),
		uintptr(unsafe.Pointer(&tinfo)))
	if hr != 0 {
		err = I创建错误对象(hr)
	}
	return
}

func invoke(disp *IDispatch, dispid int32, dispatch int16, params ...interface{}) (result *I变体结构, err error) {
	var dispparams DISPPARAMS

	if dispatch&DISPATCH_PROPERTYPUT != 0 {
		dispnames := [1]int32{DISPID_PROPERTYPUT}
		dispparams.rgdispidNamedArgs = uintptr(unsafe.Pointer(&dispnames[0]))
		dispparams.cNamedArgs = 1
	} else if dispatch&DISPATCH_PROPERTYPUTREF != 0 {
		dispnames := [1]int32{DISPID_PROPERTYPUT}
		dispparams.rgdispidNamedArgs = uintptr(unsafe.Pointer(&dispnames[0]))
		dispparams.cNamedArgs = 1
	}
	var vargs []I变体结构
	if len(params) > 0 {
		vargs = make([]I变体结构, len(params))
		for i, v := range params {
			//n := len(params)-i-1
			n := len(params) - i - 1
			I变体对象重置(&vargs[n])
			switch vv := v.(type) {
			case bool:
				if vv {
					vargs[n] = I创建变体对象(VT_BOOL, 0xffff)
				} else {
					vargs[n] = I创建变体对象(VT_BOOL, 0)
				}
			case *bool:
				vargs[n] = I创建变体对象(VT_BOOL|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*bool)))))
			case uint8:
				vargs[n] = I创建变体对象(VT_I1, int64(v.(uint8)))
			case *uint8:
				vargs[n] = I创建变体对象(VT_I1|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*uint8)))))
			case int8:
				vargs[n] = I创建变体对象(VT_I1, int64(v.(int8)))
			case *int8:
				vargs[n] = I创建变体对象(VT_I1|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*int8)))))
			case int16:
				vargs[n] = I创建变体对象(VT_I2, int64(v.(int16)))
			case *int16:
				vargs[n] = I创建变体对象(VT_I2|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*int16)))))
			case uint16:
				vargs[n] = I创建变体对象(VT_UI2, int64(v.(uint16)))
			case *uint16:
				vargs[n] = I创建变体对象(VT_UI2|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*uint16)))))
			case int32:
				vargs[n] = I创建变体对象(VT_I4, int64(v.(int32)))
			case *int32:
				vargs[n] = I创建变体对象(VT_I4|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*int32)))))
			case uint32:
				vargs[n] = I创建变体对象(VT_UI4, int64(v.(uint32)))
			case *uint32:
				vargs[n] = I创建变体对象(VT_UI4|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*uint32)))))
			case int64:
				vargs[n] = I创建变体对象(VT_I8, int64(v.(int64)))
			case *int64:
				vargs[n] = I创建变体对象(VT_I8|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*int64)))))
			case uint64:
				vargs[n] = I创建变体对象(VT_UI8, int64(uintptr(v.(uint64))))
			case *uint64:
				vargs[n] = I创建变体对象(VT_UI8|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*uint64)))))
			case int:
				vargs[n] = I创建变体对象(VT_I4, int64(v.(int)))
			case *int:
				vargs[n] = I创建变体对象(VT_I4|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*int)))))
			case uint:
				vargs[n] = I创建变体对象(VT_UI4, int64(v.(uint)))
			case *uint:
				vargs[n] = I创建变体对象(VT_UI4|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*uint)))))
			case float32:
				vargs[n] = I创建变体对象(VT_R4, *(*int64)(unsafe.Pointer(&vv)))
			case *float32:
				vargs[n] = I创建变体对象(VT_R4|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*float32)))))
			case float64:
				vargs[n] = I创建变体对象(VT_R8, *(*int64)(unsafe.Pointer(&vv)))
			case *float64:
				vargs[n] = I创建变体对象(VT_R8|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*float64)))))
			case *big.Int:
				vargs[n] = I创建变体对象(VT_DECIMAL, v.(*big.Int).Int64())
			case string:
				vargs[n] = I创建变体对象(VT_BSTR, int64(uintptr(unsafe.Pointer(SysAllocStringLen(v.(string))))))
			case *string:
				vargs[n] = I创建变体对象(VT_BSTR|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*string)))))
			case time.Time:
				s := vv.Format("2006-01-02 15:04:05")
				vargs[n] = I创建变体对象(VT_BSTR, int64(uintptr(unsafe.Pointer(SysAllocStringLen(s)))))
			case *time.Time:
				s := vv.Format("2006-01-02 15:04:05")
				vargs[n] = I创建变体对象(VT_BSTR|VT_BYREF, int64(uintptr(unsafe.Pointer(&s))))
			case *IDispatch:
				vargs[n] = I创建变体对象(VT_DISPATCH, int64(uintptr(unsafe.Pointer(v.(*IDispatch)))))
			case **IDispatch:
				vargs[n] = I创建变体对象(VT_DISPATCH|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(**IDispatch)))))
			case nil:
				vargs[n] = I创建变体对象(VT_NULL, 0)
			case *I变体结构:
				vargs[n] = I创建变体对象(VT_VARIANT|VT_BYREF, int64(uintptr(unsafe.Pointer(v.(*I变体结构)))))
			case []byte:
				safeByteArray := safeArrayFromByteSlice(v.([]byte))
				vargs[n] = I创建变体对象(VT_ARRAY|VT_UI1, int64(uintptr(unsafe.Pointer(safeByteArray))))
				defer I变体对象释放(&vargs[n])
			case []string:
				safeByteArray := safeArrayFromStringSlice(v.([]string))
				vargs[n] = I创建变体对象(VT_ARRAY|VT_BSTR, int64(uintptr(unsafe.Pointer(safeByteArray))))
				defer I变体对象释放(&vargs[n])
			default:
				panic("unknown type")
			}
		}
		dispparams.rgvarg = uintptr(unsafe.Pointer(&vargs[0]))
		dispparams.cArgs = uint32(len(params))
	}

	result = new(I变体结构)
	var excepInfo EXCEPINFO
	I变体对象重置(result)
	hr, _, _ := syscall.Syscall9(
		disp.VTable().Invoke,
		9,
		uintptr(unsafe.Pointer(disp)),
		uintptr(dispid),
		uintptr(unsafe.Pointer(IID_NULL)),
		uintptr(I取用户默认LCID()),
		uintptr(dispatch),
		uintptr(unsafe.Pointer(&dispparams)),
		uintptr(unsafe.Pointer(result)),
		uintptr(unsafe.Pointer(&excepInfo)),
		0)
	if hr != 0 {
		excepInfo.renderStrings()
		excepInfo.Clear()
		err = I创建错误对象_带子错误(hr, excepInfo.description, excepInfo)
	}
	for i, varg := range vargs {
		n := len(params) - i - 1
		if varg.VT == VT_BSTR && varg.Val != 0 {
			SysFreeString((*int16)(unsafe.Pointer(uintptr(varg.Val))))
		}
		if varg.VT == (VT_BSTR|VT_BYREF) && varg.Val != 0 {
			*(params[n].(*string)) = Unicode指针到文本(*(**uint16)(unsafe.Pointer(uintptr(varg.Val))))
		}
	}
	return
}
