func (a *{{.Obj}}) {{.FuncName}}(
	{{- range $i, $it := .Arg -}}
	{{- if gt $i 0}}, {{end}}a{{$i}} {{if $it.Variable}}...{{end}}{{$it.Type -}}
	{{- end -}}
) {{if ne .RetType ""}}{{.RetType}} {{end}}{
	{{range $i, $it := .Arg -}}
	{{if and $it.Variable $it.UserObj -}}
	av := make([]interface{}, len(a{{$i}}))
	for i, it := range a{{$i}} {
		av[i] = it.Obj
	}
	{{end -}}
	{{end -}}
	{{if ne .RetObj "" -}}
	return &{{.RetObj}}{
		{{.BasicObj}}: a.Merge(a.Obj.{{.OleFunc}}("{{.OleName}}"
		{{- range $i, $it := .Arg}}, a{{if and $it.Variable $it.UserObj}}v{{else}}{{$i}}{{end}}{{if $it.Variable}}...{{end}}{{end -}}
		)),
	}
	{{else -}}
	v, err := a.Obj.{{.OleFunc}}("{{.OleName}}"
	{{- range $i, $it := .Arg}}, a{{if and $it.Variable $it.UserObj}}v{{else}}{{$i}}{{end}}{{if $it.Variable}}...{{end}}{{end -}}
	)
	a.Merge(v, err)
	{{if ne .RetType "" -}}
	return {{if eq .RetType "string" -}}
	ToString(v, err)
	{{else if eq .RetType "bool" -}}
	ToBool(v, err)
	{{else if eq .RetType "time.Time" -}}
	ToTime(v, err)
	{{else if eq .RetType "*ole.VARIANT" -}}
	v
	{{else -}}
	({{.RetType}})(v.Val)
	{{end -}}
	{{end -}}
	{{end -}}
}
