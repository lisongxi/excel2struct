package excel2struct

import (
	"context"
	"fmt"
	"math"
	"os"
	"strconv"
	"testing"
	"time"

	"github.com/zeebo/assert"
)

type FileStruct struct {
	ID         int64     `gorm:"id"`
	Name       string    `gorm:"name" excel:"name,required"`
	Age        int8      `gorm:"age" excel:"age,required"`
	Address    string    `gorm:"address" excel:"address"`
	Birthday   time.Time `gorm:"birthday" excel:"birthday,required"`
	Height     float64   `gorm:"height" excel:"height,required"`
	IsStaff    bool      `gorm:"id_staff" excel:"isStaff,required"`
	Speed      int16     `gorm:"speed" excel:"speed"`
	Hobby      string    `gorm:"hobby" excel:"爱好"`
	WhatTime   int64     `gorm:"what_time" excel:"whatTime" parser:"unixNano"`
	CreateTime int64     `gorm:"create_time"`
}

type FileWithStruct struct {
	ID         int64     `gorm:"id"`
	Name       string    `gorm:"name" excel:"name,required"`
	Age        int8      `gorm:"age" excel:"age,required" default:"18"`
	Address    string    `gorm:"address" excel:"address"`
	Birthday   time.Time `gorm:"birthday" excel:"birthday,required"`
	Height     float64   `gorm:"height" excel:"height,required" parser:"myheight"`
	IsStaff    bool      `gorm:"id_staff" excel:"isStaff,required"`
	Speed      int16     `gorm:"speed" excel:"speed"`
	Hobby      string    `gorm:"hobby" excel:"爱好"`
	WhatTime   int64     `gorm:"what_time" excel:"whatTime" parser:"unixNano"`
	CreateTime int64     `gorm:"create_time"`
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

	var fileStruct []*FileStruct

	err = excelParser.Reader(ctx, file, &fileStruct, true)
	assert.Nil(t, err)

	for _, fs := range fileStruct {
		fmt.Printf("%+v\n", fs)
	}
}

func TestReaderWith(t *testing.T) {
	ctx := context.Background()

	file, err := os.Open("testdata/test1.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer file.Close()

	opts := []Option{
		WithFieldParser("myheight", func(field string) (interface{}, error) {
			if len(field) == 0 {
				return 0.00, nil
			}
			f64, err := strconv.ParseFloat(field, 32)
			if err != nil {
				return int64(0), err
			}
			return 2 * math.Round(f64*100) / 100, nil
		}),
	}

	excelParser, err := NewExcelParser("test1.xlsx", 0, "Sheet1", opts...)
	assert.Nil(t, err)

	fileStruct := []*FileWithStruct{}

	err = excelParser.Reader(ctx, file, &fileStruct, false)
	assert.Nil(t, err)
}

func TestReaderWorkers(t *testing.T) {
	ctx := context.Background()

	file, err := os.Open("testdata/test1.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer file.Close()

	opt := WithWorkers(12)
	excelParser, err := NewExcelParser("test1.xlsx", 0, "Sheet1", opt)
	assert.Nil(t, err)

	fileStruct := []*FileStruct{}

	err = excelParser.Reader(ctx, file, &fileStruct, false)
	assert.Nil(t, err)
}
