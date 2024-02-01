// 汇编

package dmsoft

func (com *Dmsoft) AsmAdd(asmIns string) int {
	ret, _ := com.dm.I调用方法("AsmAdd", asmIns)
	return int(ret.Val)
}

func (com *Dmsoft) AsmCall(hwnd, mode int) int {
	ret, _ := com.dm.I调用方法("AsmCall", hwnd, mode)
	return int(ret.Val)
}

func (com *Dmsoft) AsmCallEx(hwnd, mode int, baseAddr string) int64 {
	ret, _ := com.dm.I调用方法("AsmCallEx", hwnd, mode, baseAddr)
	return ret.Val
}

func (com *Dmsoft) AsmClear() int {
	ret, _ := com.dm.I调用方法("AsmClear")
	return int(ret.Val)
}

func (com *Dmsoft) AsmSetTimeout(timeOut, param int) int {
	ret, _ := com.dm.I调用方法("AsmSetTimeout", timeOut, param)
	return int(ret.Val)
}

func (com *Dmsoft) Assemble(timeOut int64, is64bit int) string {
	ret, _ := com.dm.I调用方法("Assemble", timeOut, is64bit)
	return ret.I取文本()
}

func (com *Dmsoft) DisAssemble(asmCode string, baseAddr int64, is64bit int) string {
	ret, _ := com.dm.I调用方法("DisAssemble", asmCode, baseAddr, is64bit)
	return ret.I取文本()
}
