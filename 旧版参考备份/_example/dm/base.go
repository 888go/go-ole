package dmsoft

func (com *Dmsoft) EnablePicCache(enable int) int {
	ret, _ := com.dm.I调用方法("EnablePicCache", enable)
	return int(ret.Val)
}

func (com *Dmsoft) GetBasePath() string {
	ret, _ := com.dm.I调用方法("GetBasePath")
	return ret.I取文本()
}

func (com *Dmsoft) GetDmCount() int {
	ret, _ := com.dm.I调用方法("GetDmCount")
	return int(ret.Val)
}

func (com *Dmsoft) GetID() int {
	ret, _ := com.dm.I调用方法("GetID")
	return int(ret.Val)
}

func (com *Dmsoft) GetLastError() int {
	ret, _ := com.dm.I调用方法("GetLastError")
	return int(ret.Val)
}

func (com *Dmsoft) GetPath() string {
	ret, _ := com.dm.I调用方法("GetPath")
	return ret.I取文本()
}

func (com *Dmsoft) Reg(regCode string, verInfo string) int {
	ret, _ := com.dm.I调用方法("Reg", regCode, verInfo)
	return int(ret.Val)
}

func (com *Dmsoft) RegEx(regCode, verInfo, ip string) int {
	ret, _ := com.dm.I调用方法("RegEx", regCode, verInfo, ip)
	return int(ret.Val)
}

func (com *Dmsoft) RegExNoMac(regCode, verInfo, ip string) int {
	ret, _ := com.dm.I调用方法("RegExNoMac", regCode, verInfo, ip)
	return int(ret.Val)
}

func (com *Dmsoft) RegNoMac(regCode, verInfo, ip string) int {
	ret, _ := com.dm.I调用方法("RegNoMac", regCode, verInfo, ip)
	return int(ret.Val)
}

func (com *Dmsoft) SetDisplayInput(mode string) int {
	ret, _ := com.dm.I调用方法("SetDisplayInput", mode)
	return int(ret.Val)
}

func (com *Dmsoft) SetEnumWindowDelay(delay int) int {
	ret, _ := com.dm.I调用方法("SetEnumWindowDelay", delay)
	return int(ret.Val)
}

func (com *Dmsoft) SetPath(path string) int {
	ret, _ := com.dm.I调用方法("SetPath", path)
	return int(ret.Val)
}

func (com *Dmsoft) SetShowErrorMsg(show int) int {
	ret, _ := com.dm.I调用方法("SetShowErrorMsg", show)
	return int(ret.Val)
}

func (com *Dmsoft) SpeedNormalGraphic(enable int) int {
	ret, _ := com.dm.I调用方法("SpeedNormalGraphic", enable)
	return int(ret.Val)
}

// Ver get version
func (com *Dmsoft) Ver() string {
	ver, _ := com.dm.I调用方法("Ver")
	return ver.I取文本()
}
