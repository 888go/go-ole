//go:build windows
// +build windows

package main

import (
	com快捷类 "e.coding.net/gogit/go/ego/core/win_go_ole_cn/oleutil"
	"log"
	"syscall"
	"unsafe"

	ole "e.coding.net/gogit/go/ego/core/win_go_ole_cn"
)

type EventReceiver struct {
	lpVtbl *EventReceiverVtbl
	ref    int32
	host   *ole.IDispatch
}

type EventReceiverVtbl struct {
	pQueryInterface   uintptr
	pAddRef           uintptr
	pRelease          uintptr
	pGetTypeInfoCount uintptr
	pGetTypeInfo      uintptr
	pGetIDsOfNames    uintptr
	pInvoke           uintptr
}

func QueryInterface(this *ole.IUnknown, iid *ole.GUID, punk **ole.IUnknown) uint32 {
	s, _ := ole.GUID对象取标识(iid)
	*punk = nil
	if ole.I是否相等_GUID(iid, ole.IID_IUnknown) ||
		ole.I是否相等_GUID(iid, ole.IID_IDispatch) {
		AddRef(this)
		*punk = this
		return ole.S_OK
	}
	if s == "{248DD893-BB45-11CF-9ABC-0080C7E7B78D}" {
		AddRef(this)
		*punk = this
		return ole.S_OK
	}
	return ole.E_NOINTERFACE
}

func AddRef(this *ole.IUnknown) int32 {
	pthis := (*EventReceiver)(unsafe.Pointer(this))
	pthis.ref++
	return pthis.ref
}

func Release(this *ole.IUnknown) int32 {
	pthis := (*EventReceiver)(unsafe.Pointer(this))
	pthis.ref--
	return pthis.ref
}

func GetIDsOfNames(this *ole.IUnknown, iid *ole.GUID, wnames []*uint16, namelen int, lcid int, pdisp []int32) uintptr {
	for n := 0; n < namelen; n++ {
		pdisp[n] = int32(n)
	}
	return uintptr(ole.S_OK)
}

func GetTypeInfoCount(pcount *int) uintptr {
	if pcount != nil {
		*pcount = 0
	}
	return uintptr(ole.S_OK)
}

func GetTypeInfo(ptypeif *uintptr) uintptr {
	return uintptr(ole.E_NOTIMPL)
}

func Invoke(this *ole.IDispatch, dispid int, riid *ole.GUID, lcid int, flags int16, dispparams *ole.DISPPARAMS, result *ole.I变体结构, pexcepinfo *ole.EXCEPINFO, nerr *uint) uintptr {
	switch dispid {
	case 0:
		log.Println("DataArrival")
		winsock := (*EventReceiver)(unsafe.Pointer(this)).host
		var data ole.I变体结构
		ole.I变体对象重置(&data)
		com快捷类.I调用方法(winsock, "GetData", &data)
		s := string(data.I取数组对象().I取字节集())
		println()
		println(s)
		println()
	case 1:
		log.Println("Connected")
		winsock := (*EventReceiver)(unsafe.Pointer(this)).host
		com快捷类.I调用方法(winsock, "SendData", "GET / HTTP/1.0\r\n\r\n")
	case 3:
		log.Println("SendProgress")
	case 4:
		log.Println("SendComplete")
	case 5:
		log.Println("Close")
		this.I释放()
	case 6:
		log.Fatal("Error")
	default:
		log.Println(dispid)
	}
	return ole.E_NOTIMPL
}

func main() {
	ole.I初始化COM库(0)

	unknown, err := com快捷类.I创建COM对象("{248DD896-BB45-11CF-9ABC-0080C7E7B78D}")
	if err != nil {
		panic(err.Error())
	}
	winsock, _ := unknown.I查找接口(ole.IID_IDispatch)
	iid, _ := ole.I创建GUID对象_从标识("{248DD893-BB45-11CF-9ABC-0080C7E7B78D}")

	dest := &EventReceiver{}
	dest.lpVtbl = &EventReceiverVtbl{}
	dest.lpVtbl.pQueryInterface = syscall.NewCallback(QueryInterface)
	dest.lpVtbl.pAddRef = syscall.NewCallback(AddRef)
	dest.lpVtbl.pRelease = syscall.NewCallback(Release)
	dest.lpVtbl.pGetTypeInfoCount = syscall.NewCallback(GetTypeInfoCount)
	dest.lpVtbl.pGetTypeInfo = syscall.NewCallback(GetTypeInfo)
	dest.lpVtbl.pGetIDsOfNames = syscall.NewCallback(GetIDsOfNames)
	dest.lpVtbl.pInvoke = syscall.NewCallback(Invoke)
	dest.host = winsock

	com快捷类.I连接对象(winsock, iid, (*ole.IUnknown)(unsafe.Pointer(dest)))
	_, err = com快捷类.I调用方法(winsock, "Connect", "127.0.0.1", 80)
	if err != nil {
		log.Fatal(err)
	}

	var m ole.Msg
	for dest.ref != 0 {
		ole.GetMessage(&m, 0, 0, 0)
		ole.DispatchMessage(&m)
	}
}
