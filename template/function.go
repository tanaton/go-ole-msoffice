{{$this := . -}}
func (a *{{.Obj}}) {{.FuncName}}(
	{{- range $i, $it := .Arg -}}
	{{- if gt $i 0}}, {{end}}a{{$i}} {{if $it.Variable}}...{{end}}{{$it.Type -}}
	{{- end -}}
) {{if ne .RetType ""}}{{.RetType}} {{end}}{
	{{if .ArgVariable -}}
		av := make([]interface{}, 0, 4)
	{{end -}}
	{{range $i, $it := .Arg -}}
		{{if $it.Variable -}}
			{{if gt $it.Minimum 0 -}}
				if len(a{{$i}}) < {{$it.Minimum}} {
					panic("{{$this.Obj}}.{{$this.FuncName}} : a{{$i}} : number of arguments is less than {{$it.Minimum}}")
				}
			{{end -}}
			{{if gt $it.Maximum 0 -}}
				if len(a{{$i}}) > {{$it.Maximum}} {
					panic("{{$this.Obj}}.{{$this.FuncName}} : a{{$i}} : number of arguments is greater than {{$it.Maximum}}")
				}
			{{end -}}
			{{if or $this.ArgVariable $it.TypeCheck -}}
				for _, it := range a{{$i}} {
					{{if $it.TypeCheck -}}
						switch it.(type) {
						{{range $j, $tc := $it.TypeCheck -}}
						case {{$tc}}:
							{{if $this.ArgVariable -}}
								{{if ne (UserObjectDereference $tc) "" -}}
									av = append(av, (it.({{$tc}})).Obj)
								{{else -}}
									av = append(av, it)
								{{end -}}
							{{end -}}
						{{end -}}
						default:
							panic("{{$this.Obj}}.{{$this.FuncName}} : a{{$i}} : type given for the argument is different")
						}
					{{else -}}
						{{if $this.ArgVariable -}}
							{{if ne $it.UserObj "" -}}
								av = append(av, it.Obj)
							{{else -}}
								av = append(av, it)
							{{end -}}
						{{end -}}
					{{end -}}
				}
			{{end -}}
		{{else -}}
			{{if eq $it.Type "interface{}" -}}
				{{if $it.TypeCheck -}}
					switch a{{$i}}.(type) {
					{{range $tci, $tc := $it.TypeCheck -}}
					case {{$tc}}:
						{{if $this.ArgVariable -}}
							{{if ne (UserObjectDereference $tc) "" -}}
								av = append(av, (a{{$i}}.({{$tc}})).Obj)
							{{else -}}
								av = append(av, a{{$i}})
							{{end -}}
						{{end -}}
					{{end -}}
					default:
						panic("{{$this.Obj}}.{{$this.FuncName}} : a{{$i}} : type given for the argument is different")
					}
				{{end -}}
			{{else -}}
				{{if $this.ArgVariable -}}
					{{if ne $it.UserObj "" -}}
						av = append(av, a{{$i}}.Obj)
					{{else -}}
						av = append(av, a{{$i}})
					{{end -}}
				{{end -}}
			{{end -}}
		{{end -}}
	{{end -}}
	{{if ne .RetObj "" -}}
		return &{{.RetObj}}{
			{{.BasicObj}}: a.Merge(a.Obj.{{.OleFunc}}("{{.OleName}}"
			{{- if .ArgVariable -}}
				, av...
			{{- else -}}
				{{- range $i, $it := .Arg -}}
					, a{{$i}}{{if $it.Variable}}...{{end -}}
				{{- end -}}
			{{- end -}}
			)),
		}
	{{else -}}
		v, err := a.Obj.{{.OleFunc}}("{{.OleName}}"
		{{- if .ArgVariable -}}
			, av...
		{{- else -}}
			{{- range $i, $it := .Arg -}}
				, a{{$i}}{{if $it.Variable}}...{{end -}}
			{{- end -}}
		{{- end -}}
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
