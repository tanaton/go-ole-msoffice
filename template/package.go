{{- $this := . -}}
package {{.Package}}

import (
	ole "github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"time"
	"unsafe"
)

{{if .Const -}}
const (
	{{range $key, $it := .Const -}}
	{{$key}} = {{$it}}
	{{end -}}
)
{{end -}}

type IDispatcher interface {
	IDispatch() *ole.IDispatch
}

type Merger interface {
	Merge(*ole.VARIANT, error) {{.BasicObj}}
}

type Operation interface {
	IDispatcher
	Merger
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

type Error []error

func (e *Error) Error() (ret string) {
	for _, it := range *e {
		if it != nil {
			ret += it.Error()
		}
	}
	return ret
}

func MultiError(e error, es ...error) error {
	var ee Error
	if len(es) <= 0 {
		ee = Error([]error{e})
	} else {
		ee = Error(append([]error{e}, es...))
	}
	return &ee
}

type {{.BasicObj}} struct {
	Obj      *ole.IDispatch
	Err      error
	children []{{.BasicObj}}
}

func (e *{{.BasicObj}}) Merge(obj *ole.VARIANT, err error) {{.BasicObj}} {
	ce := {{.BasicObj}}{
		Obj: obj.ToIDispatch(),
		Err: err,
	}
	e.children = append(e.children, ce)
	if e.Err == nil {
		if err != nil {
			e.Err = err
		}
	} else {
		if err != nil {
			e.Err = MultiError(e.Err, err)
		}
	}
	return ce
}
func (e *{{.BasicObj}}) IDispatch() *ole.IDispatch {
	return e.Obj
}
func (e *{{.BasicObj}}) Error() (ret string) {
	if e.Err != nil {
		ret = e.Err.Error()
	}
	return
}
func (e *{{.BasicObj}}) Release() {
	if e.children != nil {
		for i, _ := range e.children {
			e.children[i].Release()
		}
		e.children = nil
	}
	if e.Obj != nil {
		e.Obj.Release()
		e.Obj = nil
	}
}

{{range $key, $it := .Child -}}
type {{$key}} struct {
	{{$this.BasicObj}}
}
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

