//go:build arm
// +build arm

package com子类

// I变体结构  VARIANT
type I变体结构 struct {
	VT         VT     //  2
	wReserved1 uint16 //  4
	wReserved2 uint16 //  6
	wReserved3 uint16 //  8
	Val        int64  // 16
}
