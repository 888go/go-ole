// 鼠键

package dmsoft

import (
	ole "e.coding.net/gogit/go/ego/core/win_go_ole_cn"
)

func (com *Dmsoft) EnableMouseAccuracy(enable int) int {
	ret, _ := com.dm.I调用方法("EnableMouseAccuracy", enable)
	return int(ret.Val)
}

func (com *Dmsoft) GetCursorPos(x, y *int) int {
	intx := ole.I创建变体对象(ole.VT_I4, int64(*x))
	inty := ole.I创建变体对象(ole.VT_I4, int64(*y))
	ret, _ := com.dm.I调用方法("GetCursorPos", &intx, &inty)
	*x = int(intx.Val)
	*y = int(inty.Val)
	intx.I释放()
	inty.I释放()
	return int(ret.Val)
}

func (com *Dmsoft) GetCursorShape() string {
	ret, _ := com.dm.I调用方法("GetCursorShape")
	return ret.I取文本()
}

func (com *Dmsoft) GetCursorShapeEx(types int) string {
	ret, _ := com.dm.I调用方法("GetCursorShapeEx", types)
	return ret.I取文本()
}

func (com *Dmsoft) GetCursorSpot() string {
	ret, _ := com.dm.I调用方法("GetCursorSpot")
	return ret.I取文本()
}

func (com *Dmsoft) GetKeyState(vkCode int) int {
	ret, _ := com.dm.I调用方法("GetKeyState", vkCode)
	return int(ret.Val)
}

func (com *Dmsoft) GetMouseSpeed() int {
	ret, _ := com.dm.I调用方法("GetMouseSpeed")
	return int(ret.Val)
}

func (com *Dmsoft) KeyDown(vkCode int) int {
	ret, _ := com.dm.I调用方法("KeyDown", vkCode)
	return int(ret.Val)
}

func (com *Dmsoft) KeyDownChar(keyStr string) int {
	ret, _ := com.dm.I调用方法("KeyDownChar", keyStr)
	return int(ret.Val)
}

func (com *Dmsoft) KeyPress(vkCode int) int {
	ret, _ := com.dm.I调用方法("KeyPress", vkCode)
	return int(ret.Val)
}

func (com *Dmsoft) KeyPressChar(keyStr string) int {
	ret, _ := com.dm.I调用方法("KeyPressChar", keyStr)
	return int(ret.Val)
}

// KeyPressStr指定的字符串序列，依次按顺序按下其中的字符
func (com *Dmsoft) KeyPressStr(keyStr string, delay int) int {
	ret, _ := com.dm.I调用方法("KeyPressStr", keyStr, delay)
	return int(ret.Val)
}

func (com *Dmsoft) KeyUp(vkCode int) int {
	ret, _ := com.dm.I调用方法("KeyUp", vkCode)
	return int(ret.Val)
}

func (com *Dmsoft) KeyUpChar(keyStr string) int {
	ret, _ := com.dm.I调用方法("KeyUpChar", keyStr)
	return int(ret.Val)
}

// LeftClick 按下鼠标左键
func (com *Dmsoft) LeftClick() int {
	ret, _ := com.dm.I调用方法("LeftClick")
	return int(ret.Val)
}

// LeftDoubleClick 双击鼠标左键
func (com *Dmsoft) LeftDoubleClick() int {
	ret, _ := com.dm.I调用方法("LeftDoubleClick")
	return int(ret.Val)
}

// LeftDown 按住鼠标左键
func (com *Dmsoft) LeftDown() int {
	ret, _ := com.dm.I调用方法("LeftDown")
	return int(ret.Val)
}

func (com *Dmsoft) LeftUp() int {
	ret, _ := com.dm.I调用方法("LeftUp")
	return int(ret.Val)
}

func (com *Dmsoft) MiddleClick() int {
	ret, _ := com.dm.I调用方法("MiddleClick")
	return int(ret.Val)
}

func (com *Dmsoft) MiddleDown() int {
	ret, _ := com.dm.I调用方法("MiddleDown")
	return int(ret.Val)
}

func (com *Dmsoft) MiddleUp() int {
	ret, _ := com.dm.I调用方法("MiddleUp")
	return int(ret.Val)
}

func (com *Dmsoft) MoveR(rx, ry int) int {
	ret, _ := com.dm.I调用方法("MoveR", rx, ry)
	return int(ret.Val)
}

func (com *Dmsoft) MoveTo(x, y int) int {
	ret, _ := com.dm.I调用方法("MoveTo", x, y)
	return int(ret.Val)
}

func (com *Dmsoft) MoveToEx(x, y, w, h int) string {
	ret, _ := com.dm.I调用方法("MoveToEx", x, y, w, h)
	return ret.I取文本()
}

func (com *Dmsoft) RightClick() int {
	ret, _ := com.dm.I调用方法("RightClick")
	return int(ret.Val)
}

func (com *Dmsoft) RightDown() int {
	ret, _ := com.dm.I调用方法("RightDown")
	return int(ret.Val)
}

func (com *Dmsoft) RightUp() int {
	ret, _ := com.dm.I调用方法("RightUp")
	return int(ret.Val)
}

func (com *Dmsoft) SetKeypadDelay(types string, delay int) int {
	ret, _ := com.dm.I调用方法("SetKeypadDelay", types, delay)
	return int(ret.Val)
}

func (com *Dmsoft) SetMouseDelay(types string, delay int) int {
	ret, _ := com.dm.I调用方法("SetMouseDelay", types, delay)
	return int(ret.Val)
}

func (com *Dmsoft) SetMouseSpeed(speed int) int {
	ret, _ := com.dm.I调用方法("SetMouseSpeed", speed)
	return int(ret.Val)
}

func (com *Dmsoft) SetSimMode(mode int) int {
	ret, _ := com.dm.I调用方法("SetSimMode", mode)
	return int(ret.Val)
}

func (com *Dmsoft) WaitKey(vkCode, timeOut int) int {
	ret, _ := com.dm.I调用方法("SetSimMode", vkCode, timeOut)
	return int(ret.Val)
}

func (com *Dmsoft) WheelDown() int {
	ret, _ := com.dm.I调用方法("WheelDown")
	return int(ret.Val)
}

func (com *Dmsoft) WheelUp() int {
	ret, _ := com.dm.I调用方法("WheelUp")
	return int(ret.Val)
}
