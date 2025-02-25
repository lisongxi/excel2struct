package excel2struct

import (
	"context"
	"errors"
	"fmt"
	"github.com/lisongxi/goutils"
	"reflect"
	"strconv"
	"strings"
)

type ExcelParser struct {
	fileName     string
	headerIndex  int
	sheetName    string
	fieldParsers map[string]FieldParser
	rowErrs      *[]ErrorInfo
	errChan      chan ErrorInfo
	workers      int
}

func NewExcelParser(fileName string, headerIndex int, sheetName string, opts ...Option) (*ExcelParser, error) {
	excelParser := &ExcelParser{
		fileName:     fileName,
		headerIndex:  headerIndex,
		sheetName:    sheetName,
		fieldParsers: DefaultFieldParserMap,
		rowErrs:      &[]ErrorInfo{},
		errChan:      make(chan ErrorInfo, 10),
	}

	for _, opt := range opts {
		if err := opt(excelParser); err != nil {
			return nil, err
		}
	}
	return excelParser, nil
}

func (ep *ExcelParser) Parse(ctx context.Context, rows [][]string, output interface{}, skip bool) (err error) {
	defer close(ep.errChan)

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

	structFieldMetaMap := make(map[string]FieldMetadata)
	for i := 0; i < structType.NumField(); i++ {
		field := structType.Field(i)
		excelTag := field.Tag.Get("excel")
		excelTags := strings.Split(excelTag, ",")
		if len(excelTags) == 0 || excelTags[0] == "-" || strings.TrimSpace(excelTags[0]) == "" {
			continue
		}

		eIndexTag := field.Tag.Get("eIndex")
		var eIndex int
		if eIndexTag != "" && eIndexTag != "-" {
			eIndex, _ = strconv.Atoi(strings.TrimSpace(eIndexTag))
		}

		required := strings.Contains(excelTag, "required")
		defaultVal := field.Tag.Get("default")

		parserTag := field.Tag.Get("parser")
		parserTags := strings.Split(parserTag, ",")
		if parserTag == "" || len(parserTags) == 0 || parserTags[0] == "-" {
			parser := structType.Field(i).Type.Name()
			fieldMetadata := FieldMetadata{
				FIndex:   i,
				FName:    field.Name,
				Excel:    strings.TrimSpace(excelTags[0]),
				EIndex:   eIndex,
				Parser:   parser,
				Required: required,
				Default:  defaultVal,
			}
			structFieldMetaMap[strings.TrimSpace(excelTags[0])] = fieldMetadata
			continue
		}
		fieldMetadata := FieldMetadata{
			FIndex:   i,
			FName:    field.Name,
			Excel:    strings.TrimSpace(excelTags[0]),
			EIndex:   eIndex,
			Parser:   strings.TrimSpace(parserTags[0]),
			Required: required,
			Default:  defaultVal,
		}
		structFieldMetaMap[strings.TrimSpace(excelTags[0])] = fieldMetadata
	}

	if ep.headerIndex >= len(rows) {
		return errors.New("error excel header index")
	}
	if ep.headerIndex == len(rows)-1 {
		return
	}

	titleMap, err := ep.parseTitle(rows[ep.headerIndex], structFieldMetaMap)
	if err != nil {
		return
	}

	rows = rows[ep.headerIndex+1:]
	ep.AppendErrors(ctx)

	if ep.workers == 0 {
		results := reflect.MakeSlice(sliceType, 0, len(rows))
		for idx, row := range rows {
			out := reflect.New(structType)
			parsedErr := ep.parseRowToStruct(ctx, idx+ep.headerIndex+2, structFieldMetaMap, row, titleMap, out, skip)
			if parsedErr != nil {
				return parsedErr
			}
			results = reflect.Append(results, out)
		}
		outputValue.Elem().Set(results)
	} else {
		ep.parseWithWorkers(ctx, structFieldMetaMap, titleMap, structType, sliceType, outputValue, rows, skip)
	}

	return
}

func (ep *ExcelParser) parseRowToStruct(ctx context.Context, rowIndex int, structFieldMetaMap map[string]FieldMetadata, row []string, titleMap map[string]int, out reflect.Value, skip bool) (err error) {
	if out.Kind() != reflect.Ptr {
		return fmt.Errorf("the slice element must be a pointer")
	}
	if !out.IsValid() {
		return fmt.Errorf("the slice element is invalid")
	}

	outElem := out.Elem()

	for excelTag, fieldMeta := range structFieldMetaMap {
		tIdx, ok := titleMap[excelTag]
		if !ok {
			return fmt.Errorf(ERROR_TYPE[ERROR_FIELD_MATCH], excelTag)
		}

		if fieldMeta.EIndex > 0 && tIdx != fieldMeta.EIndex {
			if fieldMeta.EIndex >= len(row) {
				return fmt.Errorf(ERROR_TYPE[ERROR_EINDEX_EXCEED], fieldMeta.FName)
			}
			tIdx = fieldMeta.EIndex - 1
		}

		var field string
		if tIdx < len(row) {
			field = row[tIdx]
		}
		// set default value
		if field == "" && fieldMeta.Default != "" {
			field = fieldMeta.Default
		}
		if field == "" {
			if !skip && fieldMeta.Required {
				return fmt.Errorf(ERROR_TYPE[ERROR_REQUIRED], excelTag, rowIndex)
			}
			if fieldMeta.Required {
				ep.errChan <- ErrorInfo{
					Row:       rowIndex,
					Column:    excelTag,
					ErrorCode: ERROR_REQUIRED,
					ErrorMsg:  fmt.Sprintf(ERROR_TYPE[ERROR_REQUIRED], excelTag, rowIndex),
				}
				return nil
			}
			continue
		}

		fieldParser, registered := ep.fieldParsers[fieldMeta.Parser]
		if !registered {
			return fmt.Errorf(ERROR_TYPE[ERROR_NOT_REGISTED], fieldMeta.Parser)
		}

		value, err := fieldParser(field)
		if err != nil {
			if !skip && fieldMeta.Required {
				return fmt.Errorf(ERROR_TYPE[ERROR_PARSE], fieldMeta.FName, fieldMeta.Required, err)
			}
			ep.errChan <- ErrorInfo{
				Row:       rowIndex,
				Column:    excelTag,
				ErrorCode: ERROR_PARSE,
				ErrorMsg:  fmt.Sprintf(ERROR_TYPE[ERROR_PARSE], fieldMeta.FName, fieldMeta.Required, err),
			}
			if fieldMeta.Required {
				return nil
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

func (ep *ExcelParser) parseTitle(row []string, structFieldMetaMap map[string]FieldMetadata) (map[string]int, error) {
	titleMap := make(map[string]int)

	for idx, title := range row {
		trimmedTitle := strings.TrimSpace(title)
		if _, ok := titleMap[trimmedTitle]; !ok {
			titleMap[trimmedTitle] = idx
		}
	}

	for field, fieldMeta := range structFieldMetaMap {
		if !fieldMeta.Required {
			continue
		}

		if _, found := titleMap[field]; !found {
			return nil, fmt.Errorf("required field '%s' not found in title row", field)
		}
	}

	return titleMap, nil
}

func (ep *ExcelParser) AppendErrors(ctx context.Context) {
	goutils.SafeGo(ctx, func() {
		for err := range ep.errChan {
			*ep.rowErrs = append(*ep.rowErrs, err)
		}
	})
}
