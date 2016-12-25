go-ole-msoffice
====
go-oleでMS Officeを操作しやすくするラッパー群を提供します。
## Description
[go-ole](https://github.com/go-ole/go-ole)を用いてGo言語でMS Officeを操作しやすくするラッパー群と、ラッパー群をなるべく手間をかけずに生成するツールです。  
JScriptでExcelを操作する感覚に近いものを目指しています。  
一定のルールに従ってjsonを書くことでGo言語のコードを生成します。  

## Demo
- [Excelファイルの名前の定義を削除する](https://github.com/tanaton/go-ole-msoffice/blob/master/example/excel/namedelete.go)

## Install
### Excelの場合
`go get github.com/tanaton/go-ole-msoffice/excel`

## TODO
- Outlookの予定表を扱えるようにする
- （長期的には）MSDNライブラリをクロールしてある程度自動生成する

## Licence
[MIT](https://github.com/tanaton/go-ole-msoffice/blob/master/LICENSE.txt)
