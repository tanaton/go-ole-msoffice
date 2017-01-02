package main

import (
	"fmt"
	ole "github.com/go-ole/go-ole"
	"github.com/tanaton/go-ole-msoffice/excel"
	"path/filepath"
	"strings"
)

type ExcelGraph struct {
	app *excel.Application
}

type GraphItem struct {
	x     int
	count int
	rg    *excel.Range
	leg   []string
}

// グラフに描画するためのデータを取得
func (ex *ExcelGraph) getGraphRange(sheet *excel.Worksheet) []GraphItem {
	x := 1
	arr := []GraphItem{}
	maxCol := sheet.GetCells().GetItem(1, sheet.GetColumns().GetCount()).GetEnd(excel.XlToLeft).GetColumn() + 1

	// 探索
	for {
		v := sheet.GetCells().GetItem(1, x).GetValue()
		if v.Value() == nil || sheet.Err != nil {
			break
		}

		// 列の探索
		leg := []string{}
		i := x + 1
		for ; i < maxCol; i++ {
			v := sheet.GetCells().GetItem(1, i).GetValue()
			if v.Value() == nil || sheet.Err != nil {
				break
			} else {
				leg = append(leg, v.ToString())
			}
		}

		item := GraphItem{
			x:     x,                       // データ開始位置
			count: ((i - 1) - (x + 1)) + 1, // データの列数
			rg:    sheet.GetRange(sheet.GetColumns().GetItem(x+1), sheet.GetColumns().GetItem(i-1)),
			leg:   leg,
		}
		arr = append(arr, item)
		x = i + 1
	}
	return arr
}

// シートからグラフを作る
func (ex *ExcelGraph) sheetToChart(g *excel.Chart, sheet *excel.Worksheet, secondary []int) {
	j := 1
	arr := ex.getGraphRange(sheet)
	if len(arr) <= 0 {
		fmt.Println("シートにグラフ化できるデータが無いみたい")
		return
	}

	priname := []string{}
	secname := []string{}
	sec := map[int]struct{}{}
	for _, it := range secondary {
		sec[it] = struct{}{}
	}
	// 一つのレンジにまとめる
	union := arr[0].rg
	for i := 1; i < len(arr); i++ {
		union = ex.app.Union(union, arr[i].rg)
	}

	chart := g.GetChart()
	// データの設定
	chart.SetSourceData(union, excel.XlColumns)
	// グラフの種類を設定
	chart.SetChartType(excel.XlXYScatterLinesNoMarkers)
	// 凡例の位置を修正
	legend := chart.GetLegend()
	legend.SetPosition(excel.XlLegendPositionBottom)
	// 要素の設定
	for _, it := range arr {
		xcell := sheet.GetCells().GetItem(2, it.x)
		for k := 1; k <= it.count; k++ {
			// 線ごとにX軸の設定
			sc := chart.SeriesCollection().Item(j)
			if _, ok := sec[j]; ok {
				// 2軸
				sc.SetAxisGroup(excel.XlSecondary)
				secname = append(secname, it.leg[k-1])
			} else {
				priname = append(priname, it.leg[k-1])
			}
			end := xcell.GetEnd(excel.XlDown)
			rg := sheet.GetRange(xcell, end)
			sc.SetXValues(rg)
			j++
		}
	}
	// X軸の名前を取得
	var at string
	v := sheet.GetCells().GetItem(1, 1).GetValue()
	if v.Value() != nil && sheet.Err == nil {
		at = v.ToString()
	} else {
		at = "横軸"
	}
	// グラフの軸についての設定
	ex.setGraphAxis(g, strings.Join(priname, " / "), at)
	// 指定した要素を第二軸へ移動
	if len(secondary) > 0 {
		ex.setGraphAxisSecondary(g, strings.Join(secname, " / "))
	}
}

// グラフの軸を設定
func (ex *ExcelGraph) setGraphAxis(g *excel.Chart, name, axistitle string) {
	chart := g.GetChart()
	cp := chart.Axes(excel.XlCategory, excel.XlPrimary)
	vp := chart.Axes(excel.XlValue, excel.XlPrimary)

	// X軸の目盛線の表示
	cp.SetHasMajorGridlines(true)
	cp.SetHasMinorGridlines(true)
	// Y軸の目盛線の表示
	vp.SetHasMinorGridlines(true)
	// 目盛線の位置を下に移動
	cp.SetTickLabelPosition(excel.XlTickLabelPositionLow)
	// X軸ラベルを表示
	cp.SetHasTitle(true)
	at := cp.GetAxisTitle()
	at.SetText(axistitle)
	// Y軸ラベルを表示
	vp.SetHasTitle(true)
	at = vp.GetAxisTitle()
	at.SetText(name)
}

// 指定した要素を第二軸に移動
func (ex *ExcelGraph) setGraphAxisSecondary(g *excel.Chart, name string) {
	chart := g.GetChart()
	// 2軸目のY軸ラベルを表示
	vs := chart.Axes(excel.XlValue, excel.XlSecondary)
	vs.SetHasTitle(true)
	at := vs.GetAxisTitle()
	at.SetText(name)
}

// タイトルを設定
func (ex *ExcelGraph) setGraphTitle(g *excel.Chart, title string) {
	chart := g.GetChart()
	chart.SetHasTitle(true)
	ct := chart.GetChartTitle()
	ct.SetText(title)
	ct.SetPosition(excel.XlChartElementPositionAutomatic)
	ct.SetIncludeInLayout(false) // タイトルをグラフと重ねる
}

func main() {
	// COMの初期化
	ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED|ole.COINIT_DISABLE_OLE1DDE)
	// 確実に行う必要があるため
	defer ole.CoUninitialize()

	// エクセルオブジェクトの生成
	e := excel.ThisApplication()
	if e == nil {
		return
	}
	ex := ExcelGraph{app: e}
	// 生成してる感を出すためアプリケーションを表示する
	ex.app.SetVisible(true)
	// 既存のブックの読み込み
	workbooks := ex.app.GetWorkbooks()
	rp, _ := filepath.Abs("data/data.csv")
	book := workbooks.Open(rp)
	// シートの取得
	sheets := book.GetWorksheets()
	sheet := sheets.GetItem(1)
	// 空グラフの生成
	graph := sheet.ChartObjects().Add(30, 30, 500, 300)
	graph.SetName("graph.goで生成したグラフ")

	// 逐一描画すると遅いので、適当にまとめて描画する
	ex.app.SetScreenUpdating(false)
	// シート内容をグラフに変換
	ex.sheetToChart(graph, sheet, []int{1, 2})
	ex.app.SetScreenUpdating(true)

	ex.app.SetScreenUpdating(false)
	// タイトルを設定
	ex.setGraphTitle(graph, "ぶりいくじっと")
	// グラフオブジェクトをグラフシートに移動
	chart := graph.GetChart()
	chart.Location(excel.XlLocationAsNewSheet, "グラフその1")
	ex.app.SetScreenUpdating(true)

	// ブックを保存
	ex.app.SetDisplayAlerts(false)
	wp, _ := filepath.Abs("output.xlsx")
	book.SaveAs(wp, excel.XlWorkbookDefault)
	// ブックを閉じる
	book.Close()
	// Excelを閉じる
	ex.app.Quit()
	// メモリ解放みたいな感じ
	ex.app.Release()
}
