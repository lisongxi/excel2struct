package excel2struct

const (
	ERROR_UNKNOWN      = 1000
	ERROR_REQUIRED     = 1001
	ERROR_PARSE        = 1002
	ERROR_NOT_REGISTED = 1003
	ERROR_FIELD_MATCH  = 1004
)

var ERROR_TYPE = map[int]string{
	ERROR_UNKNOWN:      "unknown error: %s",
	ERROR_REQUIRED:     "field [%s] is required, but excel data is null: Row [%d]",
	ERROR_PARSE:        "unable to parse field [%s], Required [%t], Error [%v]",
	ERROR_NOT_REGISTED: "parsing func is not registered: Parser tag [%s]",
	ERROR_FIELD_MATCH:  "no excel title matching found: Struct Field [%s]",
}

type ErrorInfo struct {
	Row       int
	Column    string
	ErrorCode int
	ErrorMsg  string
}

type FieldMetadata struct {
	FIndex    int
	FName     string
	Excel     string
	Parser    string
	Converter string
	Required  bool
	Default   string
}
