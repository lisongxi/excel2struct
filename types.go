package excel2struct

const (
	ERROR_UNKNOWN  = 1000
	ERROR_REQUIRED = 1001
	ERROR_PARSE    = 1002
	ERROR_REGISTED = 1003
)

var ERROR_TYPE = map[int]string{
	ERROR_UNKNOWN:  "unknown error: %s",
	ERROR_REQUIRED: "field [%s] required, but not found: Row index [%d], Column [%s]",
	ERROR_PARSE:    "unable to parse field [%s]: Row index [%d], Column [%s]",
	ERROR_REGISTED: "parsing func [%s] is not registered: Type tag [%s]",
}

type ErrorInfo struct {
	ErrorCode int
	ErrorMsg  string
}

type FieldMetadata struct {
	FIndex   int
	FName    string
	Excel    string
	Parser   string
	Required bool
	Default  string
}
