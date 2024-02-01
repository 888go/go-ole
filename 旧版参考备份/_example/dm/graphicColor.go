// 图色

package dmsoft

import (
	ole "e.coding.net/gogit/go/ego/core/win_go_ole_cn"
)

func (com *Dmsoft) AppendPicAddr(picInfo string, addr, size int) string {
	ret, _ := com.dm.I调用方法("AppendPicAddr", picInfo, addr, size)
	return ret.I取文本()
}

func (com *Dmsoft) Capture(x1, y1, x2, y2 int, file string) int {
	ret, _ := com.dm.I调用方法("Capture", x1, y1, x2, y2, file)
	return int(ret.Val)
}

func (com *Dmsoft) CaptureGif(x1, y1, x2, y2 int, file string, delay, time int) int {
	ret, _ := com.dm.I调用方法("CaptureGif", x1, y1, x2, y2, file, delay, time)
	return int(ret.Val)
}

func (com *Dmsoft) CaptureJpg(x1, y1, x2, y2 int, file string, quality int) int {
	ret, _ := com.dm.I调用方法("CaptureJpg", x1, y1, x2, y2, file, quality)
	return int(ret.Val)
}

func (com *Dmsoft) CapturePng(x1, y1, x2, y2 int, file string) int {
	ret, _ := com.dm.I调用方法("CapturePng", x1, y1, x2, y2, file)
	return int(ret.Val)
}

func (com *Dmsoft) CapturePre(file string) int {
	ret, _ := com.dm.I调用方法("CapturePre", file)
	return int(ret.Val)
}

func (com *Dmsoft) CmpColor(x int, y int, color string, sim float32) int {
	ret, _ := com.dm.I调用方法("CmpColor", x, y, color, sim)
	return int(ret.Val)
}

func (com *Dmsoft) EnableDisplayDebug(enableDebug int) int {
	ret, _ := com.dm.I调用方法("EnableDisplayDebug", enableDebug)
	return int(ret.Val)
}

func (com *Dmsoft) EnableFindPicMultithread(enable int) int {
	ret, _ := com.dm.I调用方法("EnableFindPicMultithread", enable)
	return int(ret.Val)
}

func (com *Dmsoft) EnableGetColorByCapture(enable int) int {
	ret, _ := com.dm.I调用方法("EnableGetColorByCapture", enable)
	return int(ret.Val)
}

func (com *Dmsoft) FindColor(x1, y1, x2, y2 int, color string, sim float32, dir int, intX *int, intY *int) int {
	x := ole.I创建变体对象(ole.VT_I4, 0)
	y := ole.I创建变体对象(ole.VT_I4, 0)
	ret, _ := com.dm.I调用方法("FindColor", x1, y1, x2, y2, color, sim, dir, &x, &y)
	*intX = int(x.Val)
	*intY = int(y.Val)
	x.I释放()
	y.I释放()
	return int(ret.Val)
}

func (com *Dmsoft) FindColorBlock(x1, y1, x2, y2 int, color string, sim float32, count, width, height int, intX, intY *int) int {
	x := ole.I创建变体对象(ole.VT_I4, 0)
	y := ole.I创建变体对象(ole.VT_I4, 0)
	ret, _ := com.dm.I调用方法("FindColorBlock", x1, y1, x2, y2, color, sim, count, width, height, &x, &y)
	*intX = int(x.Val)
	*intY = int(y.Val)
	x.I释放()
	y.I释放()
	return int(ret.Val)
}

func (com *Dmsoft) FindColorBlockEx(x1, y1, x2, y2 int, color string, sim float32, count, width, height int) string {
	ret, _ := com.dm.I调用方法("FindColorBlockEx", x1, y1, x2, y2, color, sim, count, width, height)
	return ret.I取文本()
}

func (com *Dmsoft) FindColorE(x1, y1, x2, y2 int, color string, sim float32, dir int) string {
	ret, _ := com.dm.I调用方法("FindColorE", x1, y1, x2, y2, color, sim, dir)
	return ret.I取文本()
}

func (com *Dmsoft) FindColorEx(x1, y1, x2, y2 int, color string, sim float32, dir int) string {
	ret, _ := com.dm.I调用方法("FindColorEx", x1, y1, x2, y2, color, sim, dir)
	return ret.I取文本()
}

func (com *Dmsoft) FindMulColor(x1, y1, x2, y2 int, color string, sim float32) int {
	ret, _ := com.dm.I调用方法("FindMulColor", x1, y1, x2, y2, color, sim)
	return int(ret.Val)
}

func (com *Dmsoft) FindMultiColor(x1, y1, x2, y2 int, firstColor string, offsetColor string, sim float32, dir int, intX, intY *int) int {
	x := ole.I创建变体对象(ole.VT_I4, 0)
	y := ole.I创建变体对象(ole.VT_I4, 0)
	ret, _ := com.dm.I调用方法("FindMultiColor", x1, y1, x2, y2, firstColor, offsetColor, sim, dir, &x, &y)
	*intX = int(x.Val)
	*intY = int(y.Val)
	x.I释放()
	y.I释放()
	return int(ret.Val)
}

func (com *Dmsoft) FindMultiColorE(x1, y1, x2, y2 int, firstColor string, offsetColor string, sim float32, dir int) string {
	ret, _ := com.dm.I调用方法("FindMultiColorE", x1, y1, x2, y2, firstColor, offsetColor, sim, dir)
	return ret.I取文本()
}

func (com *Dmsoft) FindMultiColorEx(x1, y1, x2, y2 int, firstColor string, offsetColor string, sim float32, dir int) string {
	ret, _ := com.dm.I调用方法("FindMultiColorEx", x1, y1, x2, y2, firstColor, offsetColor, sim, dir)
	return ret.I取文本()
}

func (com *Dmsoft) FindPic(x1, y1, x2, y2 int, picName string, deltaColor string, sim float32, dir int, intX, intY *int) int {
	x := ole.I创建变体对象(ole.VT_I4, 0)
	y := ole.I创建变体对象(ole.VT_I4, 0)
	ret, _ := com.dm.I调用方法("FindPic", x1, y1, x2, y2, picName, deltaColor, sim, dir, &x, &y)
	*intX = int(x.Val)
	*intY = int(y.Val)
	x.I释放()
	y.I释放()
	return int(ret.Val)
}

func (com *Dmsoft) FindPicE(x1, y1, x2, y2 int, picName string, deltaColor string, sim float32, dir int) string {
	ret, _ := com.dm.I调用方法("FindPicE", x1, y1, x2, y2, picName, deltaColor, sim, dir)
	return ret.I取文本()
}

func (com *Dmsoft) FindPicEx(x1, y1, x2, y2 int, picName string, deltaColor string, sim float32, dir int) string {
	ret, _ := com.dm.I调用方法("FindPicEx", x1, y1, x2, y2, picName, deltaColor, sim, dir)
	return ret.I取文本()
}

func (com *Dmsoft) FindPicExS(x1, y1, x2, y2 int, picName string, deltaColor string, sim float32, dir int) string {
	ret, _ := com.dm.I调用方法("FindPicExS", x1, y1, x2, y2, picName, deltaColor, sim, dir)
	return ret.I取文本()
}

func (com *Dmsoft) FindPicMem(x1, y1, x2, y2 int, pic_info, delta_color string, sim float32, dir int, intX, intY *int) int {
	x := ole.I创建变体对象(ole.VT_I4, 0)
	y := ole.I创建变体对象(ole.VT_I4, 0)
	ret, _ := com.dm.I调用方法("FindPicMem", x1, y1, x2, y2, pic_info, delta_color, sim, dir, &x, &y)
	*intX = int(x.Val)
	*intY = int(y.Val)
	x.I释放()
	y.I释放()
	return int(ret.Val)
}

// func (com *Dmsoft)FindPicMemE(x1, y1, x2, y2, pic_info, delta_color,sim, dir,intX, intY)  string{}
// func (com *Dmsoft)FindPicMemEx(x1, y1, x2, y2, pic_info, delta_color,sim, dir,intX, intY)  string{}

func (com *Dmsoft) FindPicS(x1, y1, x2, y2 int, picName string, deltaColor string, sim float32, dir int, intX, intY *int) string {
	x := ole.I创建变体对象(ole.VT_I4, 0)
	y := ole.I创建变体对象(ole.VT_I4, 0)
	ret, _ := com.dm.I调用方法("FindPicS", x1, y1, x2, y2, picName, deltaColor, sim, dir, &x, &y)
	*intX = int(x.Val)
	*intY = int(y.Val)
	x.I释放()
	y.I释放()
	return ret.I取文本()
}

// func (com *Dmsoft)FindShape(x1, y1, x2, y2, offset_color,sim, dir,intX,intY) int{}
// func (com *Dmsoft)FindShapeE(x1, y1, x2, y2, offset_color,sim, dir,intX,intY) string{}
// func (com *Dmsoft)FindShapeEx(x1, y1, x2, y2, offset_color,sim, dir,intX,intY) string{}
func (com *Dmsoft) FreePic(pic_name string) int {
	ret, _ := com.dm.I调用方法("FreePic", pic_name)
	return int(ret.Val)
}

func (com *Dmsoft) GetAveHSV(x1, y1, x2, y2 int) string {
	ret, _ := com.dm.I调用方法("GetAveHSV", x1, y1, x2, y2)
	return ret.I取文本()
}

func (com *Dmsoft) GetAveRGB(x1, y1, x2, y2 int) string {
	ret, _ := com.dm.I调用方法("GetAveRGB", x1, y1, x2, y2)
	return ret.I取文本()
}

func (com *Dmsoft) GetColor(x, y int) string {
	ret, _ := com.dm.I调用方法("GetColor", x, y)
	return ret.I取文本()
}

func (com *Dmsoft) GetColorBGR(x, y int) string {
	ret, _ := com.dm.I调用方法("GetColorBGR", x, y)
	return ret.I取文本()
}

func (com *Dmsoft) GetColorHSV(x, y int) string {
	ret, _ := com.dm.I调用方法("GetColorHSV", x, y)
	return ret.I取文本()
}

func (com *Dmsoft) GetColorNum(x1, y1, x2, y2 int, color string, sim float32) int {
	ret, _ := com.dm.I调用方法("GetColorNum", x1, y1, x2, y2, color, sim)
	return int(ret.Val)
}

func (com *Dmsoft) GetPicSize(picName string) string {
	ret, _ := com.dm.I调用方法("GetPicSize", picName)
	return ret.I取文本()
}

func (com *Dmsoft) GetScreenData(x1, y1, x2, y2 int) int {
	ret, _ := com.dm.I调用方法("GetScreenData", x1, y1, x2, y2)
	return int(ret.Val)
}

func (com *Dmsoft) GetScreenDataBmp(x1, y1, x2, y2 int, data, size *int) int {
	d := ole.I创建变体对象(ole.VT_I4, int64(*data))
	s := ole.I创建变体对象(ole.VT_I4, int64(*size))
	ret, _ := com.dm.I调用方法("GetScreenDataBmp", x1, y1, x2, y2, &d, &s)
	*data = int(d.Val)
	*size = int(s.Val)
	d.I释放()
	s.I释放()
	return int(ret.Val)
}

func (com *Dmsoft) ImageToBmp(pic_name, bmp_name string) int {
	ret, _ := com.dm.I调用方法("ImageToBmp", pic_name, bmp_name)
	return int(ret.Val)
}

func (com *Dmsoft) IsDisplayDead(x1, y1, x2, y2, time int) int {
	ret, _ := com.dm.I调用方法("IsDisplayDead", x1, y1, x2, y2, time)
	return int(ret.Val)
}

func (com *Dmsoft) LoadPic(pic_name string) int {
	ret, _ := com.dm.I调用方法("LoadPic", pic_name)
	return int(ret.Val)
}

func (com *Dmsoft) LoadPicByte(addr, size int, pic_name string) int {
	ret, _ := com.dm.I调用方法("LoadPicByte", addr, size, pic_name)
	return int(ret.Val)
}

func (com *Dmsoft) MatchPicName(picName string) string {
	ret, _ := com.dm.I调用方法("MatchPicName", picName)
	return ret.I取文本()
}

func (com *Dmsoft) SetExcludeRegion(mode int, info string) int {
	ret, _ := com.dm.I调用方法("SetExcludeRegion", mode, info)
	return int(ret.Val)
}

func (com *Dmsoft) SetPicPwd(pwd string) int {
	ret, _ := com.dm.I调用方法("SetExcludeRegion", pwd)
	return int(ret.Val)
}
func (com *Dmsoft) RGB2BGR(rgb_color string) string {
	ret, _ := com.dm.I调用方法("RGB2BGR", rgb_color)
	return ret.I取文本()
}

func (com *Dmsoft) BGR2RGB(bgr_color string) string {
	ret, _ := com.dm.I调用方法("BGR2RGB", bgr_color)
	return ret.I取文本()
}

func (com *Dmsoft) SetFindPicMultithreadCount(count int) int {
	ret, _ := com.dm.I调用方法("SetFindPicMultithreadCount", count)
	return int(ret.Val)
}

func (com *Dmsoft) SetFindPicMultithreadLimit(limit int) int {
	ret, _ := com.dm.I调用方法("SetFindPicMultithreadLimit", limit)
	return int(ret.Val)
}

// 增加了接口FindPicSim FindPicSimEx FindPicSimE和FindPicSimMem FindPicSimMemEx FindPicSimMemE
