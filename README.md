go-ole-msoffice
====
go-oleでMS Officeを操作しやすくするラッパー群を提供します。
## 内容
[go-ole](https://github.com/go-ole/go-ole)を用いてGo言語でMS Officeを操作しやすくするラッパー群と、ラッパー群をなるべく手間をかけずに生成するツールです。  
JScriptでOfficeを操作する感覚に近いものを目指しています。  
一定のルールに従ってjsonを書くことでGo言語のコードを生成します。  

## インストール
### Excelを使いたい場合
`go get github.com/tanaton/go-ole-msoffice/excel`

### Outlookを使いたい場合
`go get github.com/tanaton/go-ole-msoffice/outlook`

## ラッパーの使い方
### 基本的な考え
- プロパティを読み出す際は「\[オブジェクト\].Get\[プロパティ名\]()」を呼び出す。
- プロパティを設定する際は「\[オブジェクト\].Set\[プロパティ名\]()」を呼び出す。
- メソッドを呼ぶ際は「\[オブジェクト\].\[メソッド名\]()」を呼び出す。
- プロパティ、メソッド名はVBAのMSDNを参照。

### サンプルコード
```go
package main

import (
	ole "github.com/go-ole/go-ole"
	"github.com/tanaton/go-ole-msoffice/excel"
)

func main() {
	// COMの初期化（go-oleで必要。必須！）
	ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED|ole.COINIT_DISABLE_OLE1DDE)
	defer ole.CoUninitialize()

	// エクセルオブジェクトの生成、取得
	e := excel.ThisApplication()
	if e == nil {
		return
	}
	defer e.Release()

	// プロパティの設定
	e.SetVisible(true)
	// プロパティの読み出し
	workbooks := e.GetWorkbooks()
	// メソッドの呼び出し
	workbooks.Open(`C:\test.xlsx`)
}
```

### デモ
- [Excelファイルの名前の定義を削除する](https://github.com/tanaton/go-ole-msoffice/blob/master/example/excel/namedelete/namedelete.go)
- [Excelでグラフを描画する](https://github.com/tanaton/go-ole-msoffice/blob/master/example/excel/graph/graph.go)
- [Outlookの予定表の情報を取得する](https://github.com/tanaton/go-ole-msoffice/blob/master/example/outlook/calendar_read/calendar_read.go)

## ラッパー生成ツールの使い方
- (後で書く。)

## 方針
- 既定プロパティ機能を利用しない。何をやっているのか分からなくなるので。

## 今後やりたいこと
- MSDNライブラリをクロールしてある程度自動生成する

## ライセンス
[MIT](https://github.com/tanaton/go-ole-msoffice/blob/master/LICENSE.txt)
