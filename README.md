go-ole-msoffice
====
go-oleでMS Officeを操作しやすくするラッパー群を提供します。
## 内容
[go-ole](https://github.com/go-ole/go-ole)を用いてGo言語でMS Officeを操作しやすくするラッパー群と、ラッパー群をなるべく手間をかけずに生成するツールです。  
JScriptでExcelを操作する感覚に近いものを目指しています。  
一定のルールに従ってjsonを書くことでGo言語のコードを生成します。  

## デモ
- [Excelファイルの名前の定義を削除する](https://github.com/tanaton/go-ole-msoffice/blob/master/example/excel/namedelete.go)
- [Outlookの予定表の情報を取得する](https://github.com/tanaton/go-ole-msoffice/blob/master/example/outlook/calendar_read.go)

## インストール
### Excelの場合
`go get github.com/tanaton/go-ole-msoffice/excel`

### Outlookの場合
`go get github.com/tanaton/go-ole-msoffice/outlook`

## 方針
- 既定プロパティ機能を利用しない。何をやっているのか分からなくなるので。

## やること
- （長期的には）MSDNライブラリをクロールしてある程度自動生成する

## ライセンス
[MIT](https://github.com/tanaton/go-ole-msoffice/blob/master/LICENSE.txt)
