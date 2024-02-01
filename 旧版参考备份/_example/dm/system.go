// 系统

package dmsoft

func (com *Dmsoft) Beep(f, duration int) int {
	ret, _ := com.dm.I调用方法("Beep", f, duration)
	return int(ret.Val)
}

func (com *Dmsoft) CheckFontSmooth() int {
	ret, _ := com.dm.I调用方法("CheckFontSmooth")
	return int(ret.Val)
}

func (com *Dmsoft) CheckUAC() int {
	ret, _ := com.dm.I调用方法("CheckUAC")
	return int(ret.Val)
}

// Delay 函数不推荐使用
// 建议time.Sleep()
func (com *Dmsoft) Delay(mis int) int {
	ret, _ := com.dm.I调用方法("Delay")
	return int(ret.Val)
}
func (com *Dmsoft) Delays(misMin, misMax int) int {
	ret, _ := com.dm.I调用方法("Delays", misMin, misMax)
	return int(ret.Val)
}

// DisableCloseDisplayAndSleep 不支持xp系统
func (com *Dmsoft) DisableCloseDisplayAndSleep() int {
	ret, _ := com.dm.I调用方法("DisableCloseDisplayAndSleep")
	return int(ret.Val)
}

func (com *Dmsoft) DisableFontSmooth() int {
	ret, _ := com.dm.I调用方法("DisableFontSmooth")
	return int(ret.Val)
}

func (com *Dmsoft) DisablePowerSave() int {
	ret, _ := com.dm.I调用方法("DisablePowerSave")
	return int(ret.Val)
}

func (com *Dmsoft) DisableScreenSave() int {
	ret, _ := com.dm.I调用方法("DisableScreenSave")
	return int(ret.Val)
}

func (com *Dmsoft) EnableFontSmooth() int {
	ret, _ := com.dm.I调用方法("EnableFontSmooth")
	return int(ret.Val)
}

func (com *Dmsoft) ExitOs(types int) int {
	ret, _ := com.dm.I调用方法("ExitOs", types)
	return int(ret.Val)
}

func (com *Dmsoft) GetClipboard() string {
	ret, _ := com.dm.I调用方法("GetClipboard")
	return ret.I取文本()
}

func (com *Dmsoft) GetCPUType() int {
	ret, _ := com.dm.I调用方法("GetCpuType")
	return int(ret.Val)
}

func (com *Dmsoft) GetCPUUsage() int {
	ret, _ := com.dm.I调用方法("GetCpuUsage")
	return int(ret.Val)
}

func (com *Dmsoft) GetDir(types int) string {
	ret, _ := com.dm.I调用方法("GetDir", types)
	return ret.I取文本()
}

// GetDiskModel 需要管理员权限
func (com *Dmsoft) GetDiskModel(index int) string {
	ret, _ := com.dm.I调用方法("GetDiskModel", index)
	return ret.I取文本()
}

// GetDiskReversion 需要管理员权限
func (com *Dmsoft) GetDiskReversion(index int) string {
	ret, _ := com.dm.I调用方法("GetDiskReversion", index)
	return ret.I取文本()
}

// GetDiskSerial 需要管理员权限
func (com *Dmsoft) GetDiskSerial(index int) string {
	ret, _ := com.dm.I调用方法("GetDiskSerial", index)
	return ret.I取文本()
}

func (com *Dmsoft) GetDisplayInfo() string {
	ret, _ := com.dm.I调用方法("GetDisplayInfo")
	return ret.I取文本()
}

func (com *Dmsoft) GetDPI() int {
	ret, _ := com.dm.I调用方法("GetDPI")
	return int(ret.Val)
}

func (com *Dmsoft) GetLocale() int {
	ret, _ := com.dm.I调用方法("GetLocale")
	return int(ret.Val)
}

func (com *Dmsoft) GetMachineCode() string {
	ret, _ := com.dm.I调用方法("GetMachineCode")
	return ret.I取文本()
}

// GetMachineCodeNoMac 需要管理员权限
func (com *Dmsoft) GetMachineCodeNoMac() string {
	ret, _ := com.dm.I调用方法("GetMachineCodeNoMac")
	return ret.I取文本()
}

func (com *Dmsoft) GetMemoryUsage() int {
	ret, _ := com.dm.I调用方法("GetMemoryUsage")
	return int(ret.Val)
}

func (com *Dmsoft) GetNetTime() string {
	ret, _ := com.dm.I调用方法("GetNetTime")
	return ret.I取文本()
}

func (com *Dmsoft) GetNetTimeByIP(ip string) string {
	ret, _ := com.dm.I调用方法("GetNetTimeByIp", ip)
	return ret.I取文本()
}

func (com *Dmsoft) GetOsBuildNumber() int {
	ret, _ := com.dm.I调用方法("GetOsBuildNumber")
	return int(ret.Val)
}

func (com *Dmsoft) GetOsType() int {
	ret, _ := com.dm.I调用方法("GetOsType")
	return int(ret.Val)
}

func (com *Dmsoft) GetScreenDepth() int {
	ret, _ := com.dm.I调用方法("GetScreenDepth")
	return int(ret.Val)
}

func (com *Dmsoft) GetScreenHeight() int {
	ret, _ := com.dm.I调用方法("GetScreenHeight")
	return int(ret.Val)
}

func (com *Dmsoft) GetScreenWidth() int {
	ret, _ := com.dm.I调用方法("GetScreenWidth")
	return int(ret.Val)
}

func (com *Dmsoft) GetTime() int {
	ret, _ := com.dm.I调用方法("GetTime")
	return int(ret.Val)
}

func (com *Dmsoft) Is64Bit() int {
	ret, _ := com.dm.I调用方法("Is64Bit")
	return int(ret.Val)
}

func (com *Dmsoft) IsSurrpotVt() int {
	ret, _ := com.dm.I调用方法("IsSurrpotVt")
	return int(ret.Val)
}

func (com *Dmsoft) Play(mediaFile string) int {
	ret, _ := com.dm.I调用方法("Play", mediaFile)
	return int(ret.Val)
}

func (com *Dmsoft) RunApp(appPath string, mode int) int {
	ret, _ := com.dm.I调用方法("RunApp", appPath, mode)
	return int(ret.Val)
}

func (com *Dmsoft) SetClipboard(value string) int {
	ret, _ := com.dm.I调用方法("SetClipboard", value)
	return int(ret.Val)
}

func (com *Dmsoft) SetDisplayAcceler(level int) int {
	ret, _ := com.dm.I调用方法("SetDisplayAcceler", level)
	return int(ret.Val)
}

func (com *Dmsoft) SetLocale() int {
	ret, _ := com.dm.I调用方法("SetLocale")
	return int(ret.Val)
}

func (com *Dmsoft) SetScreen(width, height, depth int) int {
	ret, _ := com.dm.I调用方法("SetScreen", width, height, depth)
	return int(ret.Val)
}

func (com *Dmsoft) SetUAC(enable int) int {
	ret, _ := com.dm.I调用方法("SetUAC", enable)
	return int(ret.Val)
}

func (com *Dmsoft) ShowTaskBarIcon(hwnd, isShow int) int {
	ret, _ := com.dm.I调用方法("ShowTaskBarIcon", hwnd, isShow)
	return int(ret.Val)
}

func (com *Dmsoft) Stop(id int) int {
	ret, _ := com.dm.I调用方法("Stop", id)
	return int(ret.Val)
}
