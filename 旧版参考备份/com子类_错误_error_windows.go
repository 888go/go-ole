//go:build windows
// +build windows

package com子类

import (
	"fmt"
	"syscall"
	"unicode/utf16"
)

// I错误码取文本 errstr 将错误代码转换为字符串。
func I错误码取文本(errno int) string {
	// 向windows询问剩余错误
	var flags uint32 = syscall.FORMAT_MESSAGE_FROM_SYSTEM | syscall.FORMAT_MESSAGE_ARGUMENT_ARRAY | syscall.FORMAT_MESSAGE_IGNORE_INSERTS
	b := make([]uint16, 300)
	n, err := syscall.FormatMessage(flags, 0, uint32(errno), 0, b, nil)
	if err != nil {
		return fmt.Sprintf("错误 %d (FormatMessage失败，错误为: %v)", errno, err)
	}
	// trim terminating \r and \n
	for ; n > 0 && (b[n-1] == '\n' || b[n-1] == '\r'); n-- {
	}
	return string(utf16.Decode(b[:n]))
}
