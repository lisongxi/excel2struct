package excel2struct

import (
	"context"
	"encoding/csv"
	"fmt"
	"io"
	"path/filepath"
	"reflect"
	"strings"

	"github.com/xuri/excelize/v2"
)

type ExcelParser struct {
	fileName             string
	headerIndex          int
	sheetName            string
	length               int
	fieldParsers         map[string]FieldParser
	timeLayoutParsers    map[string]TimeLayoutParser
	timeLocLayoutParsers map[string]TimeLocLayoutParser
}

func NewExcelParser(fileName string, headerIndex int, sheetName string, length int, opts ...Option) (*ExcelParser, error) {
	excelParser := &ExcelParser{
		fileName:             fileName,
		headerIndex:          headerIndex,
		sheetName:            sheetName,
		length:               length,
		fieldParsers:         DefaultFieldParserMap,
		timeLayoutParsers:    TimeLayoutParserMap,
		timeLocLayoutParsers: TimeLocLayoutParserMap,
	}
	for _, opt := range opts {
		if err := opt(excelParser); err != nil {
			return nil, err
		}
	}
	return excelParser, nil
}

func (ep *ExcelParser) Reader(ctx context.Context, reader io.Reader, output interface{}) (err error) {
	var rowData [][]string
	ext := filepath.Ext(ep.fileName)
	switch ext {
	case ".xlsx":
		rowData, err = ep.ReadXlsxFromReader(reader, ep.sheetName)
		if err != nil {
			return err
		}
	case ".csv":
		rowData, err = ep.ReadCsvFromReader(reader, ep.sheetName)
		if err != nil {
			return err
		}
	}
	if len(rowData) == 0 {
		return nil
	}
	return ep.Parse(ctx, rowData, output)
}

func (ep *ExcelParser) Parse(ctx context.Context, rows [][]string, output interface{}) (err error) {
	outputValue := reflect.ValueOf(output)
	outputType := reflect.TypeOf(output)
	for outputType.Kind() != reflect.Ptr || outputType.Elem().Kind() != reflect.Slice {
		return fmt.Errorf("output must be a slice pointer")
	}

	sliceType := outputType.Elem()
	elemType := sliceType.Elem()

	if elemType.Kind() != reflect.Ptr {
		return fmt.Errorf("the elements of a slice must be pointer types")
	}

	structType := elemType.Elem()
	if structType.Kind() != reflect.Struct {
		return fmt.Errorf("the pointer of the slice element must point to a struct")
	}

	structIdxFieldMetaMap := make(map[int]FieldMetadata)
	structFieldMetaMap := make(map[string]FieldMetadata)
	for i := 0; i < structType.NumField(); i++ {
		field := structType.Field(i)
		excelTag := field.Tag.Get("excel")
		excelTags := strings.Split(excelTag, ",")
		if len(excelTags) == 0 || excelTags[0] == "-" || strings.TrimSpace(excelTags[0]) == "" {
			continue
		}
		required := strings.Contains(excelTag, "required")

		parserTag := field.Tag.Get("parser")
		parserTags := strings.Split(parserTag, ",")
		if parserTag == "" || len(parserTags) == 0 || parserTags[0] == "-" {
			parser := structType.Field(i).Type.Name()
			fieldMetadata := FieldMetadata{
				FIndex:   i,
				FName:    field.Name,
				Excel:    strings.TrimSpace(excelTags[0]),
				Parser:   parser,
				Required: required,
			}
			structIdxFieldMetaMap[i] = fieldMetadata
			structFieldMetaMap[excelTags[0]] = fieldMetadata
			continue
		}
		fieldMetadata := FieldMetadata{
			FIndex:   i,
			FName:    field.Name,
			Excel:    strings.TrimSpace(excelTags[0]),
			Parser:   strings.TrimSpace(parserTags[0]),
			Required: required,
		}
		structIdxFieldMetaMap[i] = fieldMetadata
		structFieldMetaMap[excelTags[0]] = fieldMetadata
	}
	titleMap := make(map[int]string)
	array := reflect.MakeSlice(sliceType, 0, ep.length)
	for idx, row := range rows {
		temp := row
		if idx < ep.headerIndex {
			continue
		}
		if idx == ep.headerIndex {
			titleMap, err = ep.parseTitle(temp, structFieldMetaMap)
			if err != nil {
				return err
			}
			continue
		}

		out := reflect.New(structType)
		parsedErr := ep.parseRowToStruct(ctx, idx, structFieldMetaMap, temp, titleMap, out)
		if parsedErr != nil {
			continue
		}
		array = reflect.Append(array, out)
	}
	outputValue.Elem().Set(array)
	return
}

func (ep *ExcelParser) parseRowToStruct(ctx context.Context, rowIndex int, structFieldMetaMap map[string]FieldMetadata, row []string, titleMap map[int]string, out reflect.Value) (err error) {
	if out.Kind() != reflect.Ptr {
		return fmt.Errorf("the out must be a pointer")
	}
	if !out.IsValid() {
		return fmt.Errorf("the out is invalid")
	}

	outElem := out.Elem()

	for colIdx, field := range row {
		field = strings.TrimSpace(field)
		title, ok := titleMap[colIdx]
		if !ok {
			continue
		}

		fieldMeta, ok := structFieldMetaMap[title]
		if !ok {
			continue
		}

		fieldParser, registered := ep.fieldParsers[fieldMeta.Parser]
		if !registered {
			return fmt.Errorf("field parser [%s] not registered", fieldMeta.Parser)
		}

		value, err := fieldParser(field)
		if err != nil {
			if fieldMeta.Required {
				return fmt.Errorf("failed to parse required field %s (row %d, col %d): %v", fieldMeta.FName, rowIndex, colIdx, err)
			}
			continue
		}

		thisField := outElem.Field(fieldMeta.FIndex)
		if !thisField.CanSet() && !thisField.IsValid() {
			return fmt.Errorf("field not found in struct or cannot be set")
		}

		rvo := reflect.ValueOf(value)
		if thisField.Kind() == reflect.Ptr && rvo.IsZero() {
			thisField.Set(reflect.Zero(thisField.Type()))
			continue
		}

		thisField.Set(rvo)
	}

	return nil
}

func (ep *ExcelParser) ReadXlsxFromReader(reader io.Reader, sheetName string) ([][]string, error) {
	file, err := excelize.OpenReader(reader)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	if sheetName == "" {
		sheetName = file.GetSheetName(1)
	}

	rows, err := file.GetRows(sheetName)
	if err != nil {
		return nil, err
	}

	return rows, nil
}

func (ep *ExcelParser) ReadCsvFromReader(reader io.Reader, sheetName string) ([][]string, error) {
	csvReader := csv.NewReader(reader)
	csvReader.LazyQuotes = true    // 允许宽松的引号规则
	csvReader.FieldsPerRecord = -1 // 禁用字段数量验证
	rows, err := csvReader.ReadAll()
	if err != nil {
		return nil, err
	}
	return rows, nil
}

func (ep *ExcelParser) parseTitle(row []string, structFieldMetaMap map[string]FieldMetadata) (map[int]string, error) {
	fieldMap := make(map[int]string)
	titleSet := make(map[string]struct{})

	for idx, title := range row {
		trimmedTitle := strings.TrimSpace(title)
		fieldMap[idx] = trimmedTitle
		titleSet[trimmedTitle] = struct{}{}
	}

	for field, fieldMeta := range structFieldMetaMap {
		if !fieldMeta.Required {
			continue
		}

		if _, found := titleSet[field]; !found {
			return nil, fmt.Errorf("required field '%s' not found in title row", field)
		}
	}

	return fieldMap, nil
}
