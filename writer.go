package excel2struct

import (
	"context"

	"github.com/xuri/excelize/v2"
)

func (sc *StructConverter) Writer(ctx context.Context, structLists interface{}) (err error) {
	f := excelize.NewFile()
	defer f.Close()
	if sc.sheetName == "" {
		sc.sheetName = "Sheet1"
	}
	streamWriter, err := f.NewStreamWriter(sc.sheetName)
	if err != nil {
		return
	}

	err = sc.Converter(ctx, streamWriter, structLists)
	if err != nil {
		return
	}

	err = streamWriter.Flush()
	if err != nil {
		return
	}

	err = f.SaveAs(sc.filePath + sc.fileName)
	if err != nil {
		return
	}

	return
}
