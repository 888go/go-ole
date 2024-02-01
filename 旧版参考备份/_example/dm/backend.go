// 后台

package dmsoft

func (com *Dmsoft) BindWindow(hwnd int, display string, mouse string, keypad string, mode int) int {
	ret, _ := com.dm.I调用方法("BindWindow", hwnd, display, mouse, keypad, mode)
	return int(ret.Val)
}

func (com *Dmsoft) BindWindowEx(hwnd int, display string, mouse string, keypad string, public string, mode int) int {
	ret, _ := com.dm.I调用方法("BindWindowEx", hwnd, display, mouse, keypad, public, mode)
	return int(ret.Val)
}

func (com *Dmsoft) DownCPU(rate int) int {
	ret, _ := com.dm.I调用方法("DownCpu", rate)
	return int(ret.Val)
}

func (com *Dmsoft) EnableBind(enable int) int {
	ret, _ := com.dm.I调用方法("EnableBind", enable)
	return int(ret.Val)
}

func (com *Dmsoft) EnableFakeActive(enable int) int {
	ret, _ := com.dm.I调用方法("EnableFakeActive", enable)
	return int(ret.Val)
}

func (com *Dmsoft) EnableIme(enable int) int {
	ret, _ := com.dm.I调用方法("EnableIme", enable)
	return int(ret.Val)
}

func (com *Dmsoft) EnableKeypadMsg(enable int) int {
	ret, _ := com.dm.I调用方法("EnableKeypadMsg", enable)
	return int(ret.Val)
}

func (com *Dmsoft) EnableKeypadPatch(enable int) int {
	ret, _ := com.dm.I调用方法("EnableKeypadPatch", enable)
	return int(ret.Val)
}

func (com *Dmsoft) EnableKeypadSync(enable, timeOut int) int {
	ret, _ := com.dm.I调用方法("EnableKeypadSync", enable, timeOut)
	return int(ret.Val)
}

func (com *Dmsoft) EnableMouseMsg(enable int) int {
	ret, _ := com.dm.I调用方法("EnableMouseMsg", enable)
	return int(ret.Val)
}

func (com *Dmsoft) EnableMouseSync(enable, timeOut int) int {
	ret, _ := com.dm.I调用方法("EnableMouseSync", enable, timeOut)
	return int(ret.Val)
}

func (com *Dmsoft) EnableRealKeypad(enable int) int {
	ret, _ := com.dm.I调用方法("EnableRealKeypad", enable)
	return int(ret.Val)
}

func (com *Dmsoft) EnableRealMouse(enable, mousedelay, mousestep int) int {
	ret, _ := com.dm.I调用方法("EnableRealMouse", enable, mousedelay, mousestep)
	return int(ret.Val)
}

func (com *Dmsoft) EnableSpeedDx(enable int) int {
	ret, _ := com.dm.I调用方法("EnableSpeedDx", enable)
	return int(ret.Val)
}

func (com *Dmsoft) ForceUnBindWindow() int {
	ret, _ := com.dm.I调用方法("ForceUnBindWindow")
	return int(ret.Val)
}

func (com *Dmsoft) GetBindWindow() int {
	ret, _ := com.dm.I调用方法("GetBindWindow")
	return int(ret.Val)
}

func (com *Dmsoft) GetFps() int {
	ret, _ := com.dm.I调用方法("GetFps")
	return int(ret.Val)
}

func (com *Dmsoft) HackSpeed(rate int) int {
	ret, _ := com.dm.I调用方法("HackSpeed", rate)
	return int(ret.Val)
}

func (com *Dmsoft) IsBind(hwnd int) int {
	ret, _ := com.dm.I调用方法("IsBind", hwnd)
	return int(ret.Val)
}

func (com *Dmsoft) LockDisplay(lock int) int {
	ret, _ := com.dm.I调用方法("LockDisplay", lock)
	return int(ret.Val)
}

func (com *Dmsoft) LockInput(lock int) int {
	ret, _ := com.dm.I调用方法("LockInput", lock)
	return int(ret.Val)
}

func (com *Dmsoft) LockMouseRect(x1, y1, x2, y2 int) int {
	ret, _ := com.dm.I调用方法("LockMouseRect", x1, y1, x2, y2)
	return int(ret.Val)
}

func (com *Dmsoft) SetAero(enable int) int {
	ret, _ := com.dm.I调用方法("SetAero", enable)
	return int(ret.Val)
}

func (com *Dmsoft) SetDisplayDelay(time int) int {
	ret, _ := com.dm.I调用方法("SetDisplayDelay", time)
	return int(ret.Val)
}

func (com *Dmsoft) SetDisplayRefreshDelay(time int) int {
	ret, _ := com.dm.I调用方法("SetDisplayRefreshDelay", time)
	return int(ret.Val)
}

func (com *Dmsoft) SwitchBindWindow(hwnd int) int {
	ret, _ := com.dm.I调用方法("SwitchBindWindow", hwnd)
	return int(ret.Val)
}

func (com *Dmsoft) UnBindWindow() int {
	ret, _ := com.dm.I调用方法("UnBindWindow")
	return int(ret.Val)
}

func (com *Dmsoft) SetInputDm(dm_id, rx, ry int) int {
	ret, _ := com.dm.I调用方法("SetInputDm", dm_id, rx, ry)
	return int(ret.Val)
}
