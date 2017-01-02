package main

import (
	"fmt"
	ole "github.com/go-ole/go-ole"
	"github.com/tanaton/go-ole-msoffice/excel"
	"path/filepath"
)

func main() {
	// COMの初期化
	ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED|ole.COINIT_DISABLE_OLE1DDE)
	defer ole.CoUninitialize()

	// エクセルオブジェクトの生成
	e := excel.ThisApplication()
	if e == nil {
		return
	}
	defer e.Release()
	// 自動で動いている感を出すためアプリケーションを表示する
	e.SetVisible(true)
	// 既存のブックの読み込み
	workbooks := e.GetWorkbooks()
	rp, _ := filepath.Abs("data/namedelete_test.xlsx")
	workbooks.Open(rp)
	count := workbooks.GetCount()
	// ブックの選択
	book := workbooks.GetItem(count)
	// 名前リストの取得
	names := book.GetNames()
	count = names.GetCount()
	for i := 0; i < count; i++ {
		name := names.Item(1)
		n := name.GetName()
		if name.Err == nil {
			name.Delete()
			if name.Err == nil {
				fmt.Printf("%s => 削除\n", n)
			} else {
				fmt.Printf("%s => %s\n", n, name.Err)
			}
		} else {
			fmt.Println(name.Err)
		}
	}
	e.SetDisplayAlerts(false)
	wp, _ := filepath.Abs("data/namedelete_test_output.xlsx")
	book.SaveAs(wp, excel.XlWorkbookDefault)
	// ブックを閉じる
	book.Close()
	// Excelを閉じる
	e.Quit()
}
