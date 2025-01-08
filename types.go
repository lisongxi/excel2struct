package excel2struct

type FieldMetadata struct {
	Excel    string `json:"excel"`
	Parser   string `json:"parser"`
	Required bool   `json:"required"`
}
