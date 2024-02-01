// 文字识别

package dmsoft

import (
	ole "e.coding.net/gogit/go/ego/core/win_go_ole_cn"
)

// AddDict 给指定的字库中添加一条字库信息
func (com *Dmsoft) AddDict(index int, dictInfo string) int {
	ret, _ := com.dm.I调用方法("AddDict", index, dictInfo)
	return int(ret.Val)
}

// ClearDict 清空指定的字库
func (com *Dmsoft) ClearDict(index int) int {
	ret, _ := com.dm.I调用方法("ClearDict", index)
	return int(ret.Val)
}

func (com *Dmsoft) EnableShareDict(enable int) int {
	ret, _ := com.dm.I调用方法("EnableShareDict", enable)
	return int(ret.Val)
}

func (com *Dmsoft) FetchWord(x1, y1, x2, y2 int, color, word string) string {
	ret, _ := com.dm.I调用方法("FetchWord", x1, y1, x2, y2, color, word)
	return ret.I取文本()
}

func (com *Dmsoft) FindStr(x1, y1, x2, y2 int, str string, colorFormat string, sim float32, intX *int, intY *int) int {
	x := ole.I创建变体对象(ole.VT_I4, 0)
	y := ole.I创建变体对象(ole.VT_I4, 0)
	ret, _ := com.dm.I调用方法("FindStr", x1, y1, x2, y2, str, colorFormat, sim, &x, &y)
	*intX = int(x.Val)
	*intY = int(y.Val)
	x.I释放()
	y.I释放()
	return int(ret.Val)
}

func (com *Dmsoft) FindStrE(x1, y1, x2, y2 int, str string, colorFormat string, sim float32) string {
	ret, _ := com.dm.I调用方法("FindStrE", x1, y1, x2, y2, str, colorFormat, sim)
	return ret.I取文本()
}

func (com *Dmsoft) FindStrEx(x1, y1, x2, y2 int, str, colorFormat string, sim float32) string {
	ret, _ := com.dm.I调用方法("FindStrEx", x1, y1, x2, y2, str, colorFormat, sim)
	return ret.I取文本()
}

func (com *Dmsoft) FindStrExS(x1, y1, x2, y2 int, str, colorFormat string, sim float32) string {
	ret, _ := com.dm.I调用方法("FindStrExS", x1, y1, x2, y2, str, colorFormat, sim)
	return ret.I取文本()
}

func (com *Dmsoft) FindStrFast(x1, y1, x2, y2 int, str string, colorFormat string, sim float32, intX *int, intY *int) int {
	x := ole.I创建变体对象(ole.VT_I4, 0)
	y := ole.I创建变体对象(ole.VT_I4, 0)
	ret, _ := com.dm.I调用方法("FindStrFast", x1, y1, x2, y2, str, colorFormat, sim, &x, &y)
	*intX = int(x.Val)
	*intY = int(y.Val)
	x.I释放()
	y.I释放()
	return int(ret.Val)
}

func (com *Dmsoft) FindStrFastE(x1, y1, x2, y2 int, str, colorFormat string, sim float32) string {
	ret, _ := com.dm.I调用方法("FindStrFastE", x1, y1, x2, y2, str, colorFormat, sim)
	return ret.I取文本()
}

func (com *Dmsoft) FindStrFastEx(x1, y1, x2, y2 int, str, colorFormat string, sim float32) string {
	ret, _ := com.dm.I调用方法("FindStrFastEx", x1, y1, x2, y2, str, colorFormat, sim)
	return ret.I取文本()
}

func (com *Dmsoft) FindStrFastExS(x1, y1, x2, y2 int, str, colorFormat string, sim float32) string {
	ret, _ := com.dm.I调用方法("FindStrFastExS", x1, y1, x2, y2, str, colorFormat, sim)
	return ret.I取文本()
}

func (com *Dmsoft) FindStrFastS(x1, y1, x2, y2 int, str string, colorFormat string, sim float32, intX *int, intY *int) string {
	x := ole.I创建变体对象(ole.VT_I4, 0)
	y := ole.I创建变体对象(ole.VT_I4, 0)
	ret, _ := com.dm.I调用方法("FindStrFastS", x1, y1, x2, y2, str, colorFormat, sim, &x, &y)
	*intX = int(x.Val)
	*intY = int(y.Val)
	x.I释放()
	y.I释放()
	return ret.I取文本()
}

func (com *Dmsoft) FindStrS(x1, y1, x2, y2 int, str string, colorFormat string, sim float32, intX *int, intY *int) string {
	x := ole.I创建变体对象(ole.VT_I4, 0)
	y := ole.I创建变体对象(ole.VT_I4, 0)
	ret, _ := com.dm.I调用方法("FindStrS", x1, y1, x2, y2, str, colorFormat, sim, &x, &y)
	*intX = int(x.Val)
	*intY = int(y.Val)
	x.I释放()
	y.I释放()
	return ret.I取文本()
}

func (com *Dmsoft) FindStrWithFont(x1, y1, x2, y2 int, str string, colorFormat string, sim float32, fontName string, fontSize, flag int, intX *int, intY *int) int {
	x := ole.I创建变体对象(ole.VT_I4, 0)
	y := ole.I创建变体对象(ole.VT_I4, 0)
	ret, _ := com.dm.I调用方法("FindStrWithFont", x1, y1, x2, y2, str, colorFormat, sim, fontName, fontSize, flag, &x, &y)
	*intX = int(x.Val)
	*intY = int(y.Val)
	x.I释放()
	y.I释放()
	return int(ret.Val)
}

func (com *Dmsoft) FindStrWithFontE(x1, y1, x2, y2 int, str, colorFormat string, sim float32, fontName string, fontSize, flag int) string {
	ret, _ := com.dm.I调用方法("FindStrWithFontE", x1, y1, x2, y2, str, colorFormat, sim, fontName, fontSize, flag)
	return ret.I取文本()
}

func (com *Dmsoft) FindStrWithFontEx(x1, y1, x2, y2 int, str, colorFormat string, sim float32, fontName string, fontSize, flag int) string {
	ret, _ := com.dm.I调用方法("FindStrWithFontEx", x1, y1, x2, y2, str, colorFormat, sim, fontName, fontSize, flag)
	return ret.I取文本()
}

func (com *Dmsoft) GetDict(index, fontIndex int) string {
	ret, _ := com.dm.I调用方法("GetDict", index, fontIndex)
	return ret.I取文本()
}

func (com *Dmsoft) GetDictCount(index int) int {
	ret, _ := com.dm.I调用方法("GetDictCount", index)
	return int(ret.Val)
}

func (com *Dmsoft) GetDictInfo(str, fontName string, fontSize, flag int) string {
	ret, _ := com.dm.I调用方法("GetDictInfo", str, fontName, fontSize, flag)
	return ret.I取文本()
}

func (com *Dmsoft) GetNowDict() int {
	ret, _ := com.dm.I调用方法("GetNowDict")
	return int(ret.Val)
}

func (com *Dmsoft) GetResultCount(str string) int {
	ret, _ := com.dm.I调用方法("GetResultCount", str)
	return int(ret.Val)
}

func (com *Dmsoft) GetResultPos(str string, index int, intX, intY *int) int {
	x := ole.I创建变体对象(ole.VT_I4, 0)
	y := ole.I创建变体对象(ole.VT_I4, 0)
	ret, _ := com.dm.I调用方法("GetResultPos", str, index, &x, &y)
	*intX = int(x.Val)
	*intY = int(y.Val)
	x.I释放()
	y.I释放()
	return int(ret.Val)
}

func (com *Dmsoft) GetWordResultCount(str string) int {
	ret, _ := com.dm.I调用方法("GetWordResultCount", str)
	return int(ret.Val)
}

func (com *Dmsoft) GetWordResultPos(str string, index int, intX, intY *int) int {
	x := ole.I创建变体对象(ole.VT_I4, 0)
	y := ole.I创建变体对象(ole.VT_I4, 0)
	ret, _ := com.dm.I调用方法("GetWordResultPos", str, index, &x, &y)
	*intX = int(x.Val)
	*intY = int(y.Val)
	x.I释放()
	y.I释放()
	return int(ret.Val)
}

func (com *Dmsoft) GetWordResultStr(str string, index int) string {
	ret, _ := com.dm.I调用方法("GetWordResultCount", str, index)
	return ret.I取文本()
}

func (com *Dmsoft) GetWords(x1, y1, x2, y2 int, color string, sim float32) string {
	ret, _ := com.dm.I调用方法("GetWords", x1, y1, x2, y2, color, sim)
	return ret.I取文本()
}

func (com *Dmsoft) GetWordsNoDict(x1, y1, x2, y2 int, color string) string {
	ret, _ := com.dm.I调用方法("GetWordsNoDict", x1, y1, x2, y2, color)
	return ret.I取文本()
}

func (com *Dmsoft) Ocr(x1, y1, x2, y2 int, colorFormat string, sim float32) string {
	ret, _ := com.dm.I调用方法("Ocr", x1, y1, x2, y2, colorFormat, sim)
	return ret.I取文本()
}

func (com *Dmsoft) OcrEx(x1, y1, x2, y2 int, colorFormat string, sim float32) string {
	ret, _ := com.dm.I调用方法("OcrEx", x1, y1, x2, y2, colorFormat, sim)
	return ret.I取文本()
}

func (com *Dmsoft) OcrExOne(x1, y1, x2, y2 int, colorFormat string, sim float32) string {
	ret, _ := com.dm.I调用方法("OcrExOne", x1, y1, x2, y2, colorFormat, sim)
	return ret.I取文本()
}

func (com *Dmsoft) OcrInFile(x1, y1, x2, y2 int, picName, colorFormat string, sim float32) string {
	ret, _ := com.dm.I调用方法("OcrInFile", x1, y1, x2, y2, picName, colorFormat, sim)
	return ret.I取文本()
}

func (com *Dmsoft) SaveDict(index int, file string) int {
	ret, _ := com.dm.I调用方法("SaveDict", index, file)
	return int(ret.Val)
}

func (com *Dmsoft) SetColGapNoDict(colGap int) int {
	ret, _ := com.dm.I调用方法("SetColGapNoDict", colGap)
	return int(ret.Val)
}

func (com *Dmsoft) SetDict(index int, file string) int {
	ret, _ := com.dm.I调用方法("SetDict", index, file)
	return int(ret.Val)
}

func (com *Dmsoft) SetDictMem(index, addr, size int) int {
	ret, _ := com.dm.I调用方法("SetDictMem", index, addr, size)
	return int(ret.Val)
}

func (com *Dmsoft) SetDictPwd(pwd string) int {
	ret, _ := com.dm.I调用方法("SetDictPwd", pwd)
	return int(ret.Val)
}

func (com *Dmsoft) SetExactOcr(exactOcr int) int {
	ret, _ := com.dm.I调用方法("SetExactOcr", exactOcr)
	return int(ret.Val)
}

func (com *Dmsoft) SetMinColGap(minColGap int) int {
	ret, _ := com.dm.I调用方法("SetMinColGap", minColGap)
	return int(ret.Val)
}

func (com *Dmsoft) SetMinRowGap(minRowGap int) int {
	ret, _ := com.dm.I调用方法("SetMinRowGap", minRowGap)
	return int(ret.Val)
}

func (com *Dmsoft) SetRowGapNoDict(RowGap int) int {
	ret, _ := com.dm.I调用方法("SetRowGapNoDict", RowGap)
	return int(ret.Val)
}

func (com *Dmsoft) SetWordGap(wordGap int) int {
	ret, _ := com.dm.I调用方法("SetWordGap", wordGap)
	return int(ret.Val)
}

func (com *Dmsoft) SetWordGapNoDict(wordGap int) int {
	ret, _ := com.dm.I调用方法("SetWordGapNoDict", wordGap)
	return int(ret.Val)
}

func (com *Dmsoft) SetWordLineHeight(lineHeight int) int {
	ret, _ := com.dm.I调用方法("SetWordLineHeight", lineHeight)
	return int(ret.Val)
}

func (com *Dmsoft) SetWordLineHeightNoDict(lineHeight int) int {
	ret, _ := com.dm.I调用方法("SetWordLineHeightNoDict", lineHeight)
	return int(ret.Val)
}

func (com *Dmsoft) UseDict(index int) int {
	ret, _ := com.dm.I调用方法("UseDict", index)
	return int(ret.Val)
}
