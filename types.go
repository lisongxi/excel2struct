package excel2struct

type FieldMetadata struct {
	FIndex   int
	FName    string
	Excel    string
	Parser   string
	Required bool
	Default  string
}
