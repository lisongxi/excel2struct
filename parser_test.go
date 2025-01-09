package excel2struct

import (
	"context"
	"fmt"
	"log"
	"os"
	"testing"

	"github.com/xuri/excelize/v2"
	"github.com/zeebo/assert"
)

type FileStruct struct {
	ID         int64   `gorm:"id"`
	Name       string  `gorm:"name" excel:"name,required"`
	Age        int8    `gorm:"age" excel:"age,required" parser:"int8"`
	Address    string  `gorm:"address" excel:"address"`
	Birthday   string  `gorm:"birthday" excel:"birthday,required" parser:"time"`
	Height     float64 `gorm:"height" excel:"height,required" parser:"float64"`
	IsStaff    bool    `gorm:"id_staff" excel:"isStaff,required" parser:"boolean"`
	Speed      int16   `gorm:"speed" excel:"speed" parser:"int16"`
	Hobby      string  `gorm:"hobby" excel:"爱好"`
	CreateTime int64   `gorm:"create_time"`
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

func TestExcelize(t *testing.T) {
	// 打开本地 .xlsx 文件
	filePath := "testdata/test1.xlsx" // 替换为你的文件路径
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		log.Fatalf("无法打开文件: %v", err)
	}
	defer func() {
		// 关闭文件
		if err := f.Close(); err != nil {
			log.Fatalf("关闭文件失败: %v", err)
		}
	}()

	// 获取工作表的名称
	sheets := f.GetSheetList()
	if len(sheets) == 0 {
		log.Fatalf("文件没有工作表")
	}

	// 读取第一个工作表的数据
	sheetName := sheets[0]
	rows, err := f.GetRows(sheetName)
	if err != nil {
		log.Fatalf("读取工作表数据失败: %v", err)
	}

	// 遍历并打印每一行数据
	for i, row := range rows {
		for j, colCell := range row {
			fmt.Printf("第 %d 行, 第 %d 列: %s\n", i+1, j+1, colCell)
		}
	}
}
