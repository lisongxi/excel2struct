package excel2struct

import "runtime"

type Option func(excelParser *ExcelParser) error

func WithFieldParser(tag string, parser FieldParser) Option {
	return func(excelParser *ExcelParser) error {
		excelParser.fieldParsers[tag] = parser
		return nil
	}
}

func WithWorkers(num int) Option {
	return func(excelParser *ExcelParser) error {
		if num < 0 || num > runtime.NumCPU() {
			num = runtime.NumCPU()
		}
		excelParser.workers = num
		return nil
	}
}

type WOption func(structConverter *StructConverter) error

func WithFieldConverter(tag string, converter FieldConverter) WOption {
	return func(structConverter *StructConverter) error {
		structConverter.fieldConverters[tag] = converter
		return nil
	}
}
