package excel2struct

import (
	"context"
	"fmt"
	"os"
	"reflect"
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
	WhatTime   int64     `gorm:"what_time" excel:"whatTime" parser:"timeUN"`
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

	excelParser, err := NewExcelParser("test1.xlsx", 0, "Sheet1", 0)
	assert.Nil(t, err)

	fileStruct := []*FileStruct{}

	err = excelParser.Reader(ctx, file, &fileStruct)
	assert.Nil(t, err)

	for _, result := range fileStruct {
		fmt.Printf("%v\n", result)
	}
}

func TestAin(t *testing.T) {
	// 创建一个 time.Time 实例
	t1 := false

	// 获取 time.Time 的类型信息
	typ := reflect.TypeOf(t1)

	// 打印类型名
	fmt.Println("类型名:", typ.Name())    // 输出: Time
	fmt.Println("类型种类:", typ.Kind())   // 输出: struct
	fmt.Println("完整类型:", typ.String()) // 输出: time.Time
}
