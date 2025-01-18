package excel2struct

var DefaultFieldConverterMap map[string]FieldConverter

type FieldConverter func(field interface{}) (interface{}, error)
