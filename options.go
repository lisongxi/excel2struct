package excel2struct

import (
	"errors"
	"strings"
)

type Option func(excelParser *ExcelParser) error

func WithFieldParser(tag string, parser FieldParser) Option {
	return func(excelParser *ExcelParser) error {
		excelParser.fieldParsers[tag] = parser
		return nil
	}
}

func WithTimeLayoutParser(tag string, parser TimeLayoutParser) Option {
	return func(excelParser *ExcelParser) error {
		if !strings.HasPrefix(tag, "time") {
			return errors.New("time field name must start with 'time'")
		}
		excelParser.timeLayoutParsers[tag] = parser
		return nil
	}
}

func WithTimeLocLayoutParser(tag string, parser TimeLocLayoutParser) Option {
	return func(excelParser *ExcelParser) error {
		if !strings.HasPrefix(tag, "time") {
			return errors.New("time field name must start with 'time'")
		}
		excelParser.timeLocLayoutParsers[tag] = parser
		return nil
	}
}
