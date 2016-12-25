package main

import (
	"bytes"
	"encoding/json"
	"errors"
	"go/format"
	"io/ioutil"
	"log"
	"os"
	"path/filepath"
	"strings"
	"text/template"
)

const (
	InputPath  = "json"
	OutputPath = "."
)

type Argument struct {
	Type     string
	Variable bool   `json:",omitempty"`
	UserObj  string `json:",omitempty"`
}

type Function struct {
	Arguments []*Argument `json:",omitempty"`
	Return    string      `json:",omitempty"`
}

type PropertyFunc struct {
	Get *Function `json:",omitempty"`
	Set *Function `json:",omitempty"`
}

type FunctionMap struct {
	Property map[string]*PropertyFunc `json:",omitempty"`
	Method   map[string]*Function     `json:",omitempty"`
}

type Application struct {
	Package      string
	Object       string
	Const        map[string]int64
	RootFunction *Function
	Basic        FunctionMap
	Child        map[string]FunctionMap
}

type PackageData struct {
	buf          *bytes.Buffer
	Package      string
	Object       string
	BasicObj     string
	RootFunction *Function
	Const        map[string]int64
	Child        map[string]FunctionMap
}

type FunctionData struct {
	BasicObj string
	Obj      string
	FuncName string
	Arg      []*Argument
	OleFunc  string
	OleName  string
	RetType  string
	RetObj   string
}

var FunctionTempl = template.Must(template.ParseFiles("template/function.go"))
var PackageTempl = template.Must(template.ParseFiles("template/package.go"))
var FunctionPrefixMap = map[string]string{
	"GetProperty": "Get",
	"PutProperty": "Set",
	"CallMethod":  "",
}

func main() {
	var indir string
	var outroot string
	switch len(os.Args) {
	case 3:
		outroot = os.Args[2]
		fallthrough
	case 2:
		indir = os.Args[1]
	case 1, 0:
		indir = InputPath
		outroot = OutputPath
	default:
		log.Println("無効な引数です。")
		return
	}
	dir, err := ioutil.ReadDir(indir)
	if err != nil {
		log.Println(err)
		return
	}
	for _, it := range dir {
		name := it.Name()
		if filepath.Ext(name) == ".json" && it.IsDir() == false {
			err := Generate(filepath.Join(indir, name), outroot)
			if err != nil {
				log.Println(err)
			}
		}
	}
}

func Generate(in, outroot string) error {
	a, err := ReadJson(in)
	if err != nil {
		return err
	}
	if a.Package == "" {
		return errors.New("error:" + in + " - json => Packageが設定されてません。")
	}
	pd := PackageData{}
	pd.buf = &bytes.Buffer{}
	pd.Package = strings.ToLower(a.Package)
	pd.Object = a.Object
	pd.RootFunction = a.RootFunction
	pd.BasicObj = strings.Title(pd.Package)
	pd.Const = a.Const
	pd.Child = a.Child

	var terr error
	// パッケージ書き出し
	terr = PackageTempl.Execute(pd.buf, &pd)
	if terr != nil {
		log.Println(terr)
	}

	// 各関数の書き出し
	for name, it := range a.Basic.Property {
		if it.Get != nil {
			terr = pd.WriteFunction(it.Get, name, pd.BasicObj, "GetProperty")
			if terr != nil {
				log.Println(terr)
			}
		}
		if it.Set != nil {
			terr = pd.WriteFunction(it.Set, name, pd.BasicObj, "PutProperty")
			if terr != nil {
				log.Println(terr)
			}
		}
	}
	for name, it := range a.Basic.Method {
		terr = pd.WriteFunction(it, name, pd.BasicObj, "CallMethod")
		if terr != nil {
			log.Println(terr)
		}
	}

	for obj, c := range a.Child {
		for name, it := range c.Property {
			if it.Get != nil {
				terr = pd.WriteFunction(it.Get, name, obj, "GetProperty")
				if terr != nil {
					log.Println(terr)
				}
			}
			if it.Set != nil {
				terr = pd.WriteFunction(it.Set, name, obj, "PutProperty")
				if terr != nil {
					log.Println(terr)
				}
			}
		}
		for name, it := range c.Method {
			terr = pd.WriteFunction(it, name, obj, "CallMethod")
			if terr != nil {
				log.Println(terr)
			}
		}
	}

	// ソースコード整形
	buf, err := format.Source(pd.buf.Bytes())
	if err != nil {
		return err
	}

	// 出力ディレクトリ生成
	outdir := filepath.Join(outroot, pd.Package)
	mkdirerr := MakeOutputDir(outdir)
	if mkdirerr != nil {
		return mkdirerr
	}
	// ファイルに書き出す
	werr := ioutil.WriteFile(filepath.Join(outdir, pd.Package+".go"), buf, 0666)
	if werr != nil {
		return werr
	}
	return nil
}

func (pd *PackageData) WriteFunction(f *Function, name, o, of string) error {
	pre := FunctionPrefixMap[of]
	for _, arg := range f.Arguments {
		arg.UserObj = UserObjectDereference(arg.Type)
	}
	fd := FunctionData{
		BasicObj: pd.BasicObj,
		Obj:      o,
		FuncName: pre + name,
		Arg:      f.Arguments,
		OleFunc:  of,
		OleName:  name,
		RetType:  f.Return,
		RetObj:   UserObjectDereference(f.Return),
	}
	return FunctionTempl.Execute(pd.buf, &fd)
}

func UserObjectDereference(t string) (ret string) {
	switch t {
	case "", "byte", "int", "int16", "uint16", "int32", "uint32", "int64", "uint64", "bool", "string", "time.Time", "*ole.VARIANT":
		ret = ""
	default:
		if t[0] == '*' {
			ret = t[1:]
		} else {
			ret = ""
		}
	}
	return
}

func MakeOutputDir(outdir string) error {
	if st, err := os.Stat(outdir); err == nil {
		if st.IsDir() == false {
			return errors.New("出力ディレクトリがディレクトリではないみたい。")
		}
	} else {
		mkdirerr := os.MkdirAll(outdir, 0666)
		if mkdirerr != nil {
			return mkdirerr
		}
	}
	return nil
}

func ReadJson(in string) (a *Application, reterr error) {
	fp, err := os.Open(in)
	if err != nil {
		return nil, err
	}
	defer fp.Close()
	a = &Application{}
	reterr = json.NewDecoder(fp).Decode(a)
	return
}