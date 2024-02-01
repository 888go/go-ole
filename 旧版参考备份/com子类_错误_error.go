package com子类

// COM错误结构 存储COM错误。
type COM错误结构 struct {
	hr          uintptr
	description string
	subError    error
}

// I创建错误对象  NewError 使用HResult创建新错误。
func I创建错误对象(hr uintptr) *COM错误结构 {
	return &COM错误结构{hr: hr}
}

// I创建错误对象_带描述  NewErrorWithDescription 使用HResult和描述创建新的COM错误。
func I创建错误对象_带描述(hr uintptr, 描述 string) *COM错误结构 {
	return &COM错误结构{hr: hr, description: 描述}
}

// I创建错误对象_带子错误  NewErrorWithSubError 创建带有父错误的新COM错误。
func I创建错误对象_带子错误(hr uintptr, 描述 string, 子错误 error) *COM错误结构 {
	return &COM错误结构{hr: hr, description: 描述, subError: 子错误}
}

// I取错误码 Code
func (v *COM错误结构) I取错误码() uintptr {
	return uintptr(v.hr)
}

// I取文本 String 描述，手动设置或格式化带有错误代码的消息。
func (v *COM错误结构) I取文本() string {
	if v.description != "" {
		return I错误码取文本(int(v.hr)) + " (" + v.description + ")"
	}
	return I错误码取文本(int(v.hr))
}

// String  描述，手动设置或格式化带有错误代码的消息。
func (v *COM错误结构) String() string {
	if v.description != "" {
		return I错误码取文本(int(v.hr)) + " (" + v.description + ")"
	}
	return I错误码取文本(int(v.hr))
}

// Error 实现错误接口。
func (v *COM错误结构) Error() string {
	return v.I取文本()
}

// I取错误描述 Description 检索错误描述（如果有）。
func (v *COM错误结构) I取错误描述() string {
	return v.description
}

// I取子错误 SubError 返回父错误（如果有）。
func (v *COM错误结构) I取子错误() error {
	return v.subError
}
