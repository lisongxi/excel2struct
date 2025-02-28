package excel2struct

import (
	"context"
	"encoding/csv"
	"errors"
	"io"
	"path/filepath"

	"github.com/extrame/xls"
	"github.com/xuri/excelize/v2"
)

func (ep *ExcelParser) Reader(ctx context.Context, reader io.ReadSeeker, output interface{}, skip bool, args ...interface{}) (err error) {
	var rowData [][]string
	ext := filepath.Ext(ep.fileName)
	switch ext {
	case ".xlsx":
		rowData, err = ep.ReadXlsxFromReader(reader, ep.sheetName)
		if err != nil {
			return
		}
	case ".xls":
		e := "utf-8"
		if len(args) > 0 {
			e = args[0].(string)
		}
		rowData, err = ep.ReadXlsFromReader(reader, ep.sheetName, e)
		if err != nil {
			return
		}
	case ".csv":
		rowData, err = ep.ReadCsvFromReader(reader, ep.sheetName)
		if err != nil {
			return
		}
	default:
		rowData, err = ep.ReadXlsxFromReader(reader, ep.sheetName)
		if err != nil {
			return errors.New("unknown file type")
		}
	}
	if len(rowData) == 0 {
		return
	}
	err = ep.Parse(ctx, rowData, output, skip)
	return
}

func (ep *ExcelParser) ReadXlsxFromReader(reader io.ReadSeeker, sheetName string) ([][]string, error) {
	file, err := excelize.OpenReader(reader)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	if sheetName == "" {
		sheetName = file.GetSheetName(0)
	}

	rows, err := file.GetRows(sheetName)
	if err != nil {
		return nil, err
	}

	return rows, nil
}

func (ep *ExcelParser) ReadXlsFromReader(reader io.ReadSeeker, sheetName string, e string) ([][]string, error) {
	file, err := xls.OpenReader(reader, e)
	if err != nil {
		return nil, err
	}

	sheet := &xls.WorkSheet{}
	if sheetName == "" {
		sheet = file.GetSheet(0)
	} else {
		for sheetIndex := 0; sheetIndex < file.NumSheets(); sheetIndex++ {
			if w := file.GetSheet(sheetIndex); w.Name == sheetName {
				sheet = w
			}
		}
	}
	if sheet == nil {
		return nil, nil
	}
	fileData := make([][]string, 0)
	for i := 0; i < int(sheet.MaxRow); i++ {
		row := sheet.Row(i)
		rowData := make([]string, 0, row.LastCol())
		for col := 0; col < row.LastCol(); col++ {
			cell := row.Col(col)
			rowData = append(rowData, cell)
		}
		fileData = append(fileData, rowData)
	}

	return fileData, nil
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
