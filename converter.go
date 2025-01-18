package excel2struct

import (
	"context"
	"fmt"
	"reflect"
	"strings"

	"github.com/xuri/excelize/v2"
)

type StructConverter struct {
	fileName        string
	filePath        string
	sheetName       string
	fieldConverters map[string]FieldConverter
}

func NewStructConverter(fileName, filePath, sheetName string, opts ...WOption) *StructConverter {
	structConverter := &StructConverter{
		fileName:        fileName,
		filePath:        filePath,
		sheetName:       sheetName,
		fieldConverters: make(map[string]FieldConverter),
	}
	for _, opt := range opts {
		if err := opt(structConverter); err != nil {
			return nil
		}
	}
	return structConverter
}

func (sc *StructConverter) Converter(ctx context.Context, streamWriter *excelize.StreamWriter, structLists interface{}) (err error) {
	headers, err := sc.createHeaders(ctx, streamWriter, structLists)
	if err != nil {
		return
	}
	err = sc.writeData(ctx, headers, streamWriter, structLists)

	return
}

func (sc *StructConverter) createHeaders(ctx context.Context, streamWriter *excelize.StreamWriter, structLists interface{}) (map[int]FieldMetadata, error) {
	val := reflect.ValueOf(structLists)
	if val.Kind() != reflect.Slice {
		return nil, fmt.Errorf("the struct list data must be a slice")
	}
	if val.Len() == 0 {
		return nil, fmt.Errorf("nil struct list, no data")
	}
	elem := val.Index(0).Type()
	headers := make([]interface{}, 0, elem.NumField())
	headerSet := make(map[int]FieldMetadata)

	for i := 0; i < elem.NumField(); i++ {
		field := elem.Field(i)
		excelTag := field.Tag.Get("excel")
		excelTags := strings.Split(excelTag, ",")
		if len(excelTags) == 0 || excelTags[0] == "-" || strings.TrimSpace(excelTags[0]) == "" {
			continue
		}

		convertTag := field.Tag.Get("convert")

		headerSet[i] = FieldMetadata{
			FIndex:    i,
			Excel:     excelTags[0],
			Converter: convertTag,
		}
		headers = append(headers, excelTags[0])
	}
	//  Excel Header
	cell, err := excelize.CoordinatesToCellName(1, 1)
	if err != nil {
		return nil, err
	}
	if err = streamWriter.SetRow(cell, headers); err != nil {
		return nil, fmt.Errorf("write to header fail, err:%v", err)
	}
	return headerSet, nil
}

func (sc *StructConverter) writeData(ctx context.Context, headers map[int]FieldMetadata, streamWriter *excelize.StreamWriter, structLists interface{}) (err error) {
	listVal := reflect.ValueOf(structLists)

	for rIdx := 0; rIdx < listVal.Len(); rIdx++ {
		rowData := listVal.Index(rIdx)
		rowValues := make([]interface{}, 0, len(headers))

		for col := 0; col < rowData.NumField(); col++ {
			if fieldMeta, ok := headers[col]; ok {
				v := rowData.Field(col).Interface()
				if fieldMeta.Converter != "" {
					if converter, ok := sc.fieldConverters[fieldMeta.Converter]; ok {
						v, err = converter(v)
						if err != nil {
							return err
						}
					} else {
						return fmt.Errorf("convert func is not registered: Convert tag [%s]", fieldMeta.Converter)
					}
				}
				rowValues = append(rowValues, v)
			}
		}

		// write
		cell, err := excelize.CoordinatesToCellName(1, rIdx+2)
		if err != nil {
			return err
		}
		if err = streamWriter.SetRow(cell, rowValues); err != nil {
			return fmt.Errorf("write fail, err %v", err)
		}
	}
	return nil
}
