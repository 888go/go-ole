package com子类

import "unsafe"

// I创建变体对象  NewVariant 返回基于类型和值的新变量。
func I创建变体对象(vt VT, val int64) I变体结构 {
	return I变体结构{VT: vt, Val: val}
}

// I取COM对象  ToIUnknown 将Variant转换为Unknown对象。
func (v *I变体结构) I取IUnknown对象() *IUnknown {
	if v.VT != VT_UNKNOWN {
		return nil
	}
	return (*IUnknown)(unsafe.Pointer(uintptr(v.Val)))
}

// ToIDispatch 将变量转换为分派对象。
func (v *I变体结构) ToIDispatch() *IDispatch {
	if v.VT != VT_DISPATCH {
		return nil
	}
	return (*IDispatch)(unsafe.Pointer(uintptr(v.Val)))
}

// I取数组对象  ToArray 将变量转换为SafeArray助手。
func (v *I变体结构) I取数组对象() *SafeArrayConversion {
	if v.VT != VT_SAFEARRAY {
		if v.VT&VT_ARRAY == 0 {
			return nil
		}
	}
	var safeArray = (*SafeArray)(unsafe.Pointer(uintptr(v.Val)))
	return &SafeArrayConversion{safeArray}
}

// I取文本  ToString 将变量转换为Go字符串。
func (v *I变体结构) I取文本() string {
	if v.VT != VT_BSTR {
		return ""
	}
	return BstrToString(*(**uint16)(unsafe.Pointer(&v.Val)))
}

// I释放 变体对象的记忆。
func (v *I变体结构) I释放() error {
	return I变体对象释放(v)
}

// I取interface  Value 返回基于其类型的变量值。
//
// 当前支持的类型：2字节和4字节整数、字符串、布尔值。注意，64位整数、日期时间和其他类型存储为字符串，并将作为字符串返回。
//
// 需要进一步转换，因为这将返回interface{}。
func (v *I变体结构) I取interface() interface{} {
	switch v.VT {
	case VT_I1:
		return int8(v.Val)
	case VT_UI1:
		return uint8(v.Val)
	case VT_I2:
		return int16(v.Val)
	case VT_UI2:
		return uint16(v.Val)
	case VT_I4:
		return int32(v.Val)
	case VT_UI4:
		return uint32(v.Val)
	case VT_I8:
		return int64(v.Val)
	case VT_UI8:
		return uint64(v.Val)
	case VT_INT:
		return int(v.Val)
	case VT_UINT:
		return uint(v.Val)
	case VT_INT_PTR:
		return uintptr(v.Val) // TODO
	case VT_UINT_PTR:
		return uintptr(v.Val)
	case VT_R4:
		return *(*float32)(unsafe.Pointer(&v.Val))
	case VT_R8:
		return *(*float64)(unsafe.Pointer(&v.Val))
	case VT_BSTR:
		return v.I取文本()
	case VT_DATE:
		// VT_DATE类型将返回float64或time.time。
		d := uint64(v.Val)
		date, err := GetVariantDate(d)
		if err != nil {
			return float64(v.Val)
		}
		return date
	case VT_UNKNOWN:
		return v.I取IUnknown对象()
	case VT_DISPATCH:
		return v.ToIDispatch()
	case VT_BOOL:
		return (v.Val & 0xffff) != 0
	}
	return nil
}
