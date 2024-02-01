//go:build windows && 386
// +build windows,386

package com子类

import (
	"errors"
	"syscall"
	"time"
	"unsafe"
)

// GetVariantDate 将COM变量时间值转换为Go Time.Time。
func GetVariantDate(value uint64) (time.Time, error) {
	var st syscall.Systemtime
	v1 := uint32(value)
	v2 := uint32(value >> 32)
	r, _, _ := procVariantTimeToSystemTime.Call(uintptr(v1), uintptr(v2), uintptr(unsafe.Pointer(&st)))
	if r != 0 {
		return time.Date(int(st.Year), time.Month(st.Month), int(st.Day), int(st.Hour), int(st.Minute), int(st.Second), int(st.Milliseconds/1000), time.UTC), nil
	}
	return time.Now(), errors.New("Could not convert to time, passing current time.")
}
