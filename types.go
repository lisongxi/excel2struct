package excel2struct

const (
	ERROR_UNKNOWN       = 1000
	ERROR_REQUIRED      = 1001
	ERROR_PARSE         = 1002
	ERROR_NOT_REGISTED  = 1003
	ERROR_FIELD_MATCH   = 1004
	ERROR_EINDEX_EXCEED = 1005
)

var ERROR_TYPE = map[int]string{
	ERROR_UNKNOWN:       "unknown error: %s",
	ERROR_REQUIRED:      "field [%s] is required, but excel data is null: Row [%d]",
	ERROR_PARSE:         "unable to parse field [%s], Required [%t], Error [%v]",
	ERROR_NOT_REGISTED:  "parsing func is not registered: Parser tag [%s]",
	ERROR_FIELD_MATCH:   "no excel title matching found: Struct Field [%s]",
	ERROR_EINDEX_EXCEED: "the Excel column index settings exceed the line length, field [%s]",
}

type ErrorInfo struct {
	Row       int
	Column    string
	ErrorCode int
	ErrorMsg  string
}

type FieldMetadata struct {
	FIndex    int    // struct field index
	FName     string // struct field name
	Excel     string // excel column name
	EIndex    int    // excel column index
	Parser    string // excel column parser
	Converter string // struct to excel converter
	Required  bool
	Default   string
}
