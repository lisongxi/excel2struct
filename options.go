package excel2struct

type Option func(excelParser *ExcelParser) error

func WithFieldParser(tag string, parser FieldParser) Option {
	return func(excelParser *ExcelParser) error {
		excelParser.fieldParsers[tag] = parser
		return nil
	}
}
