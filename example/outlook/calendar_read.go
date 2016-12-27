package main

import (
	"fmt"
	ole "github.com/go-ole/go-ole"
	"github.com/tanaton/go-ole-msoffice/outlook"
	"runtime"
)

func main() {
	// スレッドを固定する
	runtime.LockOSThread()
	// スレッドの固定を解除する（※goroutineを抜けると自動で解除される）
	//defer runtime.UnlockOSThread()
	ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED|ole.COINIT_DISABLE_OLE1DDE)
	defer ole.CoUninitialize()

	ol := outlook.ThisApplication()
	defer ol.Release()

	// ネームスペース取得
	ns := ol.GetSession()

	// 予定表フォルダの取得
	folder := ns.GetDefaultFolder(outlook.OlFolderCalendar)

	items := folder.GetItems()
	count := items.GetCount()
	for i := 1; i <= count; i++ {
		item, err := outlook.AppointmentItemCast(items.Item(i))
		if err != nil {
			fmt.Println(err)
			continue
		}

		subject := item.GetSubject()
		if item.Err != nil {
			fmt.Println(item.Err)
		}
		start := item.GetStart()
		end := item.GetEnd()
		location := item.GetLocation()
		if item.Err != nil {
			fmt.Println(item.Err)
		}
		body := item.GetBody()
		if item.Err != nil {
			fmt.Println(item.Err)
		}

		// 出力
		fmt.Println("subject:", subject)
		fmt.Println("start:", start)
		fmt.Println("end:", end)
		fmt.Println("location:", location)
		fmt.Println("body:", body)
		fmt.Println("==============================================================")
	}
}
