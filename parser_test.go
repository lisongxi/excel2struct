package excel2struct

import (
	"context"
	"fmt"
	"os"
	"testing"

	"github.com/stretchr/testify/assert"
)

type FileStruct struct {
	Name     string  `excel:"name,required"`
	Age      int8    `excel:"age,required" parser:"int8"`
	Address  string  `excel:"address"`
	Birthday string  `excel:"birthday,required" parser:"time"`
	Height   float64 `excel:"height,required" parser:"float64"`
	IsStaff  bool    `excel:"isStaff,required" parser:"boolean"`
	Speed    int16   `excel:"speed" parser:"int16"`
}

func TestReader(t *testing.T) {
	ctx := context.Background()

	file, err := os.Open("testdata/test1.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer file.Close()

	excelParser, err := NewExcelParser("test1.xlsx", 0, "Sheet1")
	assert.Nil(t, err)

	fileStruct := &FileStruct{}

	results, err := excelParser.Reader(ctx, file, fileStruct)
	assert.Nil(t, err)

	for _, result := range results {
		row := result.(*FileStruct)
		fmt.Println(row)
	}
}
