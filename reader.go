package excel2struct

import (
	"context"
	"encoding/csv"
	"io"
	"path/filepath"

	"github.com/xuri/excelize/v2"
)

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
	err = ep.Parse(ctx, rowData, output)
	return
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
	csvReader.LazyQuotes = true
	csvReader.FieldsPerRecord = -1
	rows, err := csvReader.ReadAll()
	if err != nil {
		return nil, err
	}
	return rows, nil
}
