package {{.Package}}
{{$this := . -}}

import (
	"time"
	"unsafe"

	ole "github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

{{if .Const -}}
const (
	{{range $key, $it := .Const -}}
	{{$it.Name}} = {{$it.Data}}
	{{end -}}
)
{{end -}}

type IDispatcher interface {
	IDispatch() *ole.IDispatch
}

type Merger interface {
	Merge(*ole.VARIANT, error) {{.BasicObj}}
}

type Cast interface {
	Extends() {{.BasicObj}}
	GetClass() int
}

func ToString(v *ole.VARIANT, err error) (ret string) {
	if v.Value() != nil && err == nil {
		ret = v.ToString()
	}
	return
}
func ToBool(v *ole.VARIANT, err error) (ret bool) {
	if err == nil {
		if i := v.Value(); i != nil {
			if b, ok := i.(bool); ok {
				ret = b
			}
		}
	}
	return
}
func ToTime(v *ole.VARIANT, err error) (t time.Time) {
	if err == nil {
		f := *(*float64)(unsafe.Pointer(&v.Val))
		t = time.Date(1900, time.January, 1, 0, 0, 0, 0, time.Local)
		t = t.Add(time.Hour * 24 * time.Duration(int64(f)-2))
		t = t.Add(time.Millisecond * time.Duration((f-float64(int64(f)))/(1.0/86400000.0)))
	}
	return
}

type Error string

func (e Error) Error() string {
	return string(e)
}

type ErrorArray []error

func (e *ErrorArray) Error() (ret string) {
	for _, it := range *e {
		if it != nil {
			ret += it.Error()
		}
	}
	return ret
}

func MultiError(e error, es ...error) error {
	var ee ErrorArray
	if it, ok := e.(*ErrorArray); ok {
		ee = *it
	} else {
		ee = ErrorArray([]error{e})
	}
	if len(es) > 0 {
		ees := make([]error, 0, len(es)+1)
		for _, it := range es {
			if _, ok := it.(*ErrorArray); ok == false {
				ees = append(ees, it)
			}
		}
		ee = ErrorArray(append([]error(ee), ees...))
	}
	return &ee
}

type {{.BasicObj}} struct {
	Obj      *ole.IDispatch
	Err      error
	children []{{.BasicObj}}
}

func (a *{{.BasicObj}}) Merge(obj *ole.VARIANT, err error) {{.BasicObj}} {
	b := {{.BasicObj}}{
		Obj: obj.ToIDispatch(),
		Err: err,
	}
	a.children = append(a.children, b)
	if a.Err == nil {
		if err != nil {
			a.Err = err
		}
	} else {
		if err != nil {
			a.Err = MultiError(a.Err, err)
		}
	}
	return b
}
func (a *{{.BasicObj}}) Extends() {{.BasicObj}} {
	b := {{.BasicObj}}{
		Obj: a.Obj,
		Err: a.Err,
	}
	a.children = append(a.children, b)
	return b
}
func (a *{{.BasicObj}}) IDispatch() *ole.IDispatch {
	return a.Obj
}
func (a *{{.BasicObj}}) Error() (ret string) {
	if a.Err != nil {
		ret = a.Err.Error()
	}
	return
}
func (a *{{.BasicObj}}) Release() {
	if a.children != nil {
		for i, _ := range a.children {
			a.children[i].Release()
		}
		a.children = nil
	}
	if a.Obj != nil {
		a.Obj.Release()
		a.Obj = nil
	}
}

{{range $key, $it := .Child -}}
type {{$it.Objname}} struct {
	{{$this.BasicObj}}
}
{{if ne $it.TypeConst "" -}}
func {{$it.Objname}}Cast(a Cast) (*{{$it.Objname}}, error) {
	if a.GetClass() != {{$it.TypeConst}} {
		return nil, Error("Cast error : {{$it.Objname}}")
	}
	return &{{$it.Objname}}{
		{{$this.BasicObj}}: a.Extends(),
	}, nil
}
{{end -}}
{{end -}}

{{if .RootFunction -}}
func {{.RootFunction.Name}}() {{.RootFunction.Return}} {
	unknown, err := oleutil.CreateObject("{{.Object}}")
	if err != nil {
		return nil
	}
	obj, err := unknown.QueryInterface(ole.IID_IDispatch)
	return &{{.RootFunction.RetUserObj}}{
		{{.BasicObj}}: {{.BasicObj}}{Obj: obj, Err: err},
	}
}
{{end -}}

