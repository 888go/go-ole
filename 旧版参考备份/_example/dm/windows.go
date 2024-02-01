// 窗口

package dmsoft

import (
	ole "e.coding.net/gogit/go/ego/core/win_go_ole_cn"
)

func (com *Dmsoft) ClientToScreen(hwnd int, x, y *int) int {
	intx := ole.I创建变体对象(ole.VT_I4, int64(*x))
	inty := ole.I创建变体对象(ole.VT_I4, int64(*y))
	ret, _ := com.dm.I调用方法("ClientToScreen", hwnd, &intx, &inty)
	*x = int(intx.Val)
	*y = int(inty.Val)
	intx.I释放()
	inty.I释放()
	return int(ret.Val)
}

func (com *Dmsoft) EnumProcess(name string) string {
	ret, _ := com.dm.I调用方法("EnumProcess", name)
	return ret.I取文本()
}

func (com *Dmsoft) EnumWindow(parent int, title, className string, filter int) string {
	ret, _ := com.dm.I调用方法("EnumWindow", parent, title, className, filter)
	return ret.I取文本()
}

func (com *Dmsoft) EnumWindowByProcess(processName, title, className string, filter int) string {
	ret, _ := com.dm.I调用方法("EnumWindowByProcess", processName, title, className, filter)
	return ret.I取文本()
}

func (com *Dmsoft) EnumWindowByProcessId(pid int, title, className string, filter int) string {
	ret, _ := com.dm.I调用方法("EnumWindowByProcessId", pid, title, className, filter)
	return ret.I取文本()
}

func (com *Dmsoft) EnumWindowSuper(spec1 string, flag1, type1 int, spec2 string, flag2, type2, sort int) string {
	ret, _ := com.dm.I调用方法("EnumWindowSuper", spec1, flag1, type1, spec2, flag2, type2, sort)
	return ret.I取文本()
}

func (com *Dmsoft) FindWindow(class, title string) int {
	ret, _ := com.dm.I调用方法("FindWindow", class, title)
	return int(ret.Val)
}

func (com *Dmsoft) FindWindowByProcess(processName, class, title string) int {
	ret, _ := com.dm.I调用方法("FindWindowByProcess", processName, class, title)
	return int(ret.Val)
}

func (com *Dmsoft) FindWindowByProcessId(processId int, class, title string) int {
	ret, _ := com.dm.I调用方法("FindWindowByProcessId", processId, class, title)
	return int(ret.Val)
}

func (com *Dmsoft) FindWindowEx(parent int, class, title string) int {
	ret, _ := com.dm.I调用方法("FindWindowEx", parent, class, title)
	return int(ret.Val)
}

func (com *Dmsoft) FindWindowSuper(spec1 string, flag1, type1 int, spec2 string, flag2, type2 int) int {
	ret, _ := com.dm.I调用方法("FindWindowSuper", spec1, flag1, type1, spec2, flag2, type2)
	return int(ret.Val)
}

func (com *Dmsoft) GetClientRect(hwnd int, x1, y1, x2, y2 *int) int {
	intx1 := ole.I创建变体对象(ole.VT_I4, int64(*x1))
	inty1 := ole.I创建变体对象(ole.VT_I4, int64(*y1))
	intx2 := ole.I创建变体对象(ole.VT_I4, int64(*x2))
	inty2 := ole.I创建变体对象(ole.VT_I4, int64(*y2))
	ret, _ := com.dm.I调用方法("GetClientRect", hwnd, &intx1, &inty1, &intx2, &inty2)
	*x1 = int(intx1.Val)
	*y1 = int(inty1.Val)
	*x2 = int(intx2.Val)
	*y2 = int(inty2.Val)
	// 清除对象变量内存
	intx1.I释放()
	inty1.I释放()
	intx2.I释放()
	inty2.I释放()
	return int(ret.Val)
}

func (com *Dmsoft) GetClientSize(hwnd int, width, height *int) int {
	pWidth := ole.I创建变体对象(ole.VT_I4, int64(*width))
	pHeight := ole.I创建变体对象(ole.VT_I4, int64(*height))
	ret, _ := com.dm.I调用方法("GetClientSize", hwnd, &pWidth, &pHeight)
	*width = int(pWidth.Val)
	*height = int(pHeight.Val)
	pWidth.I释放()
	pHeight.I释放()
	return int(ret.Val)
}

func (com *Dmsoft) GetForegroundFocus() int {
	ret, _ := com.dm.I调用方法("GetForegroundFocus")
	return int(ret.Val)
}

func (com *Dmsoft) GetForegroundWindow() int {
	ret, _ := com.dm.I调用方法("GetForegroundWindow")
	return int(ret.Val)
}

func (com *Dmsoft) GetMousePointWindow() int {
	ret, _ := com.dm.I调用方法("GetMousePointWindow")
	return int(ret.Val)
}

func (com *Dmsoft) GetPointWindow(x, y int) int {
	ret, _ := com.dm.I调用方法("GetPointWindow", x, y)
	return int(ret.Val)
}

func (com *Dmsoft) GetProcessInfo(pid int) string {
	ret, _ := com.dm.I调用方法("GetProcessInfo", pid)
	return ret.I取文本()
}

func (com *Dmsoft) GetSpecialWindow(flag int) int {
	ret, _ := com.dm.I调用方法("GetSpecialWindow", flag)
	return int(ret.Val)
}

func (com *Dmsoft) GetWindow(hwnd, flag int) int {
	ret, _ := com.dm.I调用方法("GetWindow", hwnd, flag)
	return int(ret.Val)
}

func (com *Dmsoft) GetWindowClass(hwnd int) string {
	ret, _ := com.dm.I调用方法("GetWindowClass", hwnd)
	return ret.I取文本()
}

func (com *Dmsoft) GetWindowProcessId(hwnd int) int {
	ret, _ := com.dm.I调用方法("GetWindowProcessId", hwnd)
	return int(ret.Val)
}

func (com *Dmsoft) GetWindowProcessPath(hwnd int) string {
	ret, _ := com.dm.I调用方法("GetWindowProcessPath", hwnd)
	return ret.I取文本()
}

func (com *Dmsoft) GetWindowRect(hwnd int, x1, y1, x2, y2 *int) int {
	intx1 := ole.I创建变体对象(ole.VT_I4, int64(*x1))
	inty1 := ole.I创建变体对象(ole.VT_I4, int64(*y1))
	intx2 := ole.I创建变体对象(ole.VT_I4, int64(*x2))
	inty2 := ole.I创建变体对象(ole.VT_I4, int64(*y2))
	ret, _ := com.dm.I调用方法("GetWindowRect", hwnd, &intx1, &inty1, &intx2, &inty2)
	*x1 = int(intx1.Val)
	*y1 = int(inty1.Val)
	*x2 = int(intx2.Val)
	*y2 = int(inty2.Val)
	intx1.I释放()
	inty1.I释放()
	intx2.I释放()
	inty2.I释放()
	return int(ret.Val)
}

func (com *Dmsoft) GetWindowState(hwnd, flag int) int {
	ret, _ := com.dm.I调用方法("GetWindowState", hwnd, flag)
	return int(ret.Val)
}

func (com *Dmsoft) GetWindowTitle(hwnd int) string {
	ret, _ := com.dm.I调用方法("GetWindowTitle", hwnd)
	return ret.I取文本()
}

func (com *Dmsoft) MoveWindow(hwnd, x, y int) int {
	ret, _ := com.dm.I调用方法("MoveWindow", hwnd, x, y)
	return int(ret.Val)
}

func (com *Dmsoft) ScreenToClient(hwnd int, x, y *int) int {
	intx := ole.I创建变体对象(ole.VT_I4, int64(*x))
	inty := ole.I创建变体对象(ole.VT_I4, int64(*y))
	ret, _ := com.dm.I调用方法("ScreenToClient", hwnd, &intx, &inty)
	*x = int(intx.Val)
	*y = int(inty.Val)
	intx.I释放()
	inty.I释放()
	return int(ret.Val)
}

func (com *Dmsoft) SendPaste(hwnd int) int {
	ret, _ := com.dm.I调用方法("SendPaste", hwnd)
	return int(ret.Val)
}

func (com *Dmsoft) SendString(hwnd int, str string) int {
	ret, _ := com.dm.I调用方法("SendString", hwnd, str)
	return int(ret.Val)
}

func (com *Dmsoft) SendString2(hwnd int, str string) int {
	ret, _ := com.dm.I调用方法("SendString2", hwnd, str)
	return int(ret.Val)
}

func (com *Dmsoft) SendStringIme(str string) int {
	ret, _ := com.dm.I调用方法("SendStringIme", str)
	return int(ret.Val)
}

func (com *Dmsoft) SendStringIme2(hwnd int, str string, mode int) int {
	ret, _ := com.dm.I调用方法("SendStringIme2", hwnd, str, mode)
	return int(ret.Val)
}

func (com *Dmsoft) SetClientSize(hwnd, width, height int) int {
	ret, _ := com.dm.I调用方法("SetClientSize", hwnd, width, height)
	return int(ret.Val)
}

func (com *Dmsoft) SetWindowSize(hwnd, width, height int) int {
	ret, _ := com.dm.I调用方法("SetWindowSize", hwnd, width, height)
	return int(ret.Val)
}

func (com *Dmsoft) SetWindowState(hwnd, flag int) int {
	ret, _ := com.dm.I调用方法("SetWindowState", hwnd, flag)
	return int(ret.Val)
}

func (com *Dmsoft) SetWindowText(hwnd int, title string) int {
	ret, _ := com.dm.I调用方法("SetWindowText", hwnd, title)
	return int(ret.Val)
}

func (com *Dmsoft) SetWindowTransparent(hwnd, trans int) int {
	ret, _ := com.dm.I调用方法("SetWindowTransparent", hwnd, trans)
	return int(ret.Val)
}

func (com *Dmsoft) GetWindowThreadId(hwnd int) int {
	ret, _ := com.dm.I调用方法("GetWindowThreadId", hwnd)
	return int(ret.Val)
}
