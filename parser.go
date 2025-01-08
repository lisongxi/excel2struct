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
	fieldParsers         map[string]FieldParser
	timeLayoutParsers    map[string]TimeLayoutParser
	timeLocLayoutParsers map[string]TimeLocLayoutParser
}

func NewExcelParser(fileName string, headerIndex int, sheetName string, opts ...Option) (*ExcelParser, error) {
	excelParser := &ExcelParser{
		fileName:             fileName,
		headerIndex:          headerIndex,
		sheetName:            sheetName,
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

func (ep *ExcelParser) Reader(ctx context.Context, reader io.Reader, toStruct interface{}) (results []interface{}, err error) {
	var rowData [][]string
	ext := filepath.Ext(ep.fileName)
	switch ext {
	case ".xlsx":
		rowData, err = ep.ReadXlsxFromReader(reader, ep.sheetName)
		if err != nil {
			return nil, err
		}
	case ".csv":
		rowData, err = ep.ReadCsvFromReader(reader, ep.sheetName)
		if err != nil {
			return nil, err
		}
	}
	if len(rowData) == 0 {
		return results, nil
	}
	return ep.Parse(ctx, rowData, toStruct)
}

func (ep *ExcelParser) Parse(ctx context.Context, rows [][]string, toStruct interface{}) (results []interface{}, err error) {
	theType := reflect.TypeOf(toStruct)
	for theType.Kind() != reflect.Struct {
		if theType.Kind() == reflect.Ptr {
			theType = theType.Elem()
		} else {
			return nil, fmt.Errorf("struct required")
		}
	}
	fieldNum := theType.NumField()
	structFieldMetaMap := make(map[int]FieldMetadata)
	fieldRequiredMap := make(map[string]bool)
	for i := 0; i < fieldNum; i++ {
		field := theType.Field(i)
		excelTag := field.Tag.Get("excel")
		excelTags := strings.Split(excelTag, ",")
		if len(excelTags) == 0 || excelTags[0] == "-" || strings.TrimSpace(excelTags[0]) == "" {
			continue
		}
		required := strings.Contains(excelTag, "required")
		fieldRequiredMap[excelTags[0]] = required

		parserTag := field.Tag.Get("parser")
		parserTags := strings.Split(parserTag, ",")
		if parserTag == "" || len(parserTags) == 0 || parserTags[0] == "-" {
			parser := theType.Field(i).Type.Name()
			structFieldMetaMap[i] = FieldMetadata{
				Excel:    strings.TrimSpace(excelTags[0]),
				Parser:   parser,
				Required: required,
			}
			continue
		}
		structFieldMetaMap[i] = FieldMetadata{
			Excel:    strings.TrimSpace(excelTags[0]),
			Parser:   strings.TrimSpace(parserTags[0]),
			Required: required,
		}
	}
	titleMap := make(map[int]string)
	for idx, row := range rows {
		temp := row
		if idx < ep.headerIndex {
			continue
		}
		if idx == ep.headerIndex {
			titleMap, err = ep.parseTitle(temp, fieldRequiredMap)
			if err != nil {
				return nil, err
			}
			continue
		}

		parsedRow, parseErr := ep.parseRow(ctx, idx, structFieldMetaMap, temp, titleMap, toStruct)
		if parseErr != nil {
			return nil, parseErr
		}
		results = append(results, parsedRow)
	}
	return
}

func (ep *ExcelParser) parseRow(ctx context.Context, rowIndex int, structFieldMetaMap map[int]FieldMetadata, row []string, titleMap map[int]string, toStruct interface{}) (interface{}, error) {
	tp := reflect.TypeOf(toStruct)
	tv := reflect.New(tp.Elem())

	for colIdx, field := range row {
		field = strings.TrimSpace(field)
		if title, ok := titleMap[colIdx]; ok {
			for sIdx, fieldMeta := range structFieldMetaMap {
				fIdx := sIdx
				if fieldMeta.Excel == title {
					if fieldParser, registered := ep.fieldParsers[fieldMeta.Parser]; registered {
						value, err := fieldParser(field)
						if err != nil {
							if fieldMeta.Required {
								return nil, err
							}
							continue
						}
						tv.Elem().Field(fIdx).Set(reflect.ValueOf(value))
						break
					}
					return nil, fmt.Errorf("field parser [%s] not found", fieldMeta.Parser)
				}
			}
		}
	}
	return tv.Interface(), nil
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

func (ep *ExcelParser) parseTitle(row []string, fieldRequiredMap map[string]bool) (map[int]string, error) {
	fieldMap := make(map[int]string)
	titleSet := make(map[string]struct{})

	for idx, title := range row {
		trimmedTitle := strings.TrimSpace(title)
		fieldMap[idx] = trimmedTitle
		titleSet[trimmedTitle] = struct{}{}
	}

	for field, required := range fieldRequiredMap {
		if !required {
			continue
		}

		if _, found := titleSet[field]; !found {
			return nil, fmt.Errorf("required field '%s' not found in title row", field)
		}
	}

	return fieldMap, nil
}
