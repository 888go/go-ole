// 包旨在检索和处理从COM返回的安全数组数据。

package com子类

// SafeArrayBound 定义SafeArray边界。
type SafeArrayBound struct {
	Elements   uint32
	LowerBound int32
}

// SafeArray 是COM处理数组的方式。
type SafeArray struct {
	Dimensions   uint16
	FeaturesFlag uint16
	ElementsSize uint32
	LocksAmount  uint32
	Data         uint32
	Bounds       [16]byte
}

// SAFEARRAY 已过时，为向后兼容而存在。
// 使用SafeArray
type SAFEARRAY SafeArray

// SAFEARRAYBOUND 已过时，为向后兼容而存在。
// 使用SafeArrayBound
type SAFEARRAYBOUND SafeArrayBound
