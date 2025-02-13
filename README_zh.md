# excel2struct

## 介绍

`excel2struct`是一个实现了`struct`和`excel`相互转换的开源库，并且可以针对单个字段自定义解析方法。

## 基本使用

> `go get github.com/lisongxi/excel2struct`
### 1. 入门
| name  | age | address | birthday   | height | isStaff | speed  | 爱好   | whatTime             |
|-------|-----|---------|------------|--------|---------|--------|--------|----------------------|
| Lucas | 18  | China   | 2005/8/17  | 182.5  | T       | 10.5   | 篮球   | 2024-01-09 14:33:09  |
| John  | 25  | USA     | 1999/5/7   | 177.08 | F       | test  | 足球   | 2025-12-09 14:33:10  |
| Mike  | 33  |         | 2001/1/9   | 123    | T       |        |        | 2021-01-09 09:19:11  |
| Jeny  | 22  | Mexico  | 2024/2/3   | 456    | F       | 19     | 羽毛球 |                      |
| 小明  | 23  | 中国    | 2004/9/9   | 182.5  | F       | 12     |        |                      |

把以上数据生成一个xlsx文件，例如命名为`test.xlsx`，执行以下代码

```go
package main

import (
	"context"
	"fmt"
	"os"
	"time"

	"github.com/lisongxi/excel2struct"
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

func main() {
	ctx := context.Background()

	file, err := os.Open("testdata/test1.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer file.Close()

	excelParser, err := excel2struct.NewExcelParser("test1.xlsx", 0, "Sheet1")
	if err != nil {
		fmt.Println(err)
	}

	fileStruct := []*FileStruct{}

	err = excelParser.Reader(ctx, file, &fileStruct)
	if err != nil {
		fmt.Print(err)
	}
	for _, fs := range fileStruct {
		fmt.Printf("%+v\n", fs)
	}
}
```

如果终端能得到以下信息，说明你已经掌握`excel2struct`的基本用法了。

```shell
&{ID:0 Name:Lucas Age:18 Address:China Birthday:2005-08-17 00:00:00 +0000 UTC Height:182.5 IsStaff:true Speed:10 Hobby:篮球 WhatTime:1704810789000000000 CreateTime:0}
&{ID:0 Name:John Age:25 Address:USA Birthday:1999-05-07 00:00:00 +0000 UTC Height:177.08 IsStaff:false Speed:0 Hobby:足球 WhatTime:1765290790000000000 CreateTime:0}
&{ID:0 Name:Mike Age:33 Address: Birthday:2001-01-09 00:00:00 +0000 UTC Height:123 IsStaff:true Speed:0 Hobby: WhatTime:1610183951000000000 CreateTime:0}
&{ID:0 Name:Jeny Age:22 Address:Mexico Birthday:2024-02-03 00:00:00 +0000 UTC Height:456 IsStaff:false Speed:19 Hobby:羽毛球 WhatTime:0 CreateTime:0}
&{ID:0 Name:小明 Age:23 Address:中国 Birthday:2004-09-09 00:00:00 +0000 UTC Height:182.5 IsStaff:false Speed:12 Hobby: WhatTime:0 CreateTime:0}
```

### 2. 参数解释
```go
func NewExcelParser(fileName string, headerIndex int, sheetName string, opts ...Option) (*ExcelParser, error) 
// fileName: 文件名，必填，并且要带文件后缀，例如`.xlsx`，当前版本仅支持xlsx和csv文件;
// headerIndex：标题行所在行数据的下标索引，必填，下标从0开始，并且headerIndex以前的数据行会被忽略;
// sheetName：sheet名称，如果传入空字符串，则默认解析第一个sheet;
```

```go
func (ep *ExcelParser) Reader(ctx context.Context, reader io.Reader, output interface{}) (err error) 
// ctx：上下文
// reader: 实现了 io.Reader 接口的类型，例如打开的文件;
// output：接收excel数据的结构体指针切片，注意切片元素类型一定要是指针类型，指向结构体;
```