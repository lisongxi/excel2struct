# excel2struct

## 介绍

`excel2struct`是一个实现了`struct`和`excel`相互转换的开源库，并且可以针对单个字段自定义解析方法。

## 基本使用

> `go get github.com/lisongxi/excel2struct`
### 入门
| name  | age | address | birthday   | height | isStaff | speed  | 爱好   | whatTime             |
|-------|-----|---------|------------|--------|---------|--------|--------|----------------------|
| Lucas | 18  | China   | 2005/8/17  | 182.5  | T       | 10.5   | 篮球   | 2024-01-09 14:33:09  |
| John  | 25  | USA     | 1999/5/7   | 177.08 | F       | test  | 足球   | 2025-12-09 14:33:10  |
| Mike  | 33  |         | 2001/1/9   | 123    | T       |        |        | 2021-01-09 09:19:11  |
| Jeny  | 22  | Mexico  | 2024/2/3   | 456    | F       | 19     | 羽毛球 |                      |
| 小明  | 23  | 中国    | 2004/9/9   | 182.5  | F       | 12     |        |                      |

把以上数据生成一个xlsx文件，例如命名为`test.xlsx`，Sheet命名为`Sheet1`，执行以下代码

```go
package main

import (
	"context"
	"fmt"
	"os"
	"time"

	e2s "github.com/lisongxi/excel2struct"
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

	excelParser, err := e2s.NewExcelParser("test1.xlsx", 0, "Sheet1")
	if err != nil {
		fmt.Println(err)
	}

	fileStruct := []*FileStruct{}

	err = excelParser.Reader(ctx, file, &fileStruct, true)
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

### 参数解释
```go
func NewExcelParser(fileName string, headerIndex int, sheetName string, opts ...Option) (*ExcelParser, error) 
// fileName: 文件名，必填，并且要带文件后缀，例如`.xlsx`，当前版本仅支持xlsx、xls和csv文件，默认xlsx；
// headerIndex：标题行所在行数据的下标索引，必填，下标从0开始，并且headerIndex以前的数据行会被忽略;
// sheetName：sheet名称，如果传入空字符串，则默认解析第一个sheet;
```

```go
func (ep *ExcelParser) Reader(ctx context.Context, reader io.Reader, output interface{}, skip bool) (err error) 
// ctx：上下文
// reader: 实现了`io.Reader`接口的类型，例如打开的文件;
// output：接收excel数据的结构体指针切片，注意切片元素类型一定要是指针类型，指向结构体;
// skip: 针对`required`字段。当skip=false时，表示不允许跳过，当required字段为空且无默认值或者解析过程发生错误，则直接返回错误，不再继续往下解析；
//       当skip=true时，表示允许跳过，当required字段为空且无默认值或者解析过程发生错误，则记录错误，并跳过这一行，继续解析下一行；
```

### 标签解释
```go
`excel`: Excel文件内容的列名，如果在后面接`required`，表示该字段必填。例如`excel:"birthday,required"`；有`excel`标签才可以将`struct字段`与`Excel列`关联起来；
`parser`: 该字段对应的自定义解析函数。基本类型无需额外添加`parser`，4.2有详细说明；非必须；
`eIndex`: Excel文件内容的列索引，**从 1 开始计数**。例如`eIndex:"5"`，主要是为了解决列名有重名的情况。优先级高于`excel`标签；非必须；
```

### 错误信息
```go
excelParser, _ := e2s.NewExcelParser("test1.xlsx", 0, "Sheet1")
err = excelParser.Reader(ctx, file, &fileStruct, true)
```
- 创建的`excelParser`有结构体变量`rowErrs  *[]ErrorInfo`，它主要负责收集解析过程发生的错误；
- 错误信息定义如下：
```go
	const (
		ERROR_UNKNOWN       = 1000
		ERROR_REQUIRED      = 1001
		ERROR_PARSE         = 1002
		ERROR_NOT_REGISTED  = 1003
		ERROR_FIELD_MATCH   = 1004
		ERROR_EINDEX_EXCEED = 1005
	)

	var ERROR_TYPE = map[int]string{
		ERROR_UNKNOWN:       "unknown error: %s",
		ERROR_REQUIRED:      "field [%s] is required, but excel data is null: Row [%d]",
		ERROR_PARSE:         "unable to parse field [%s], Required [%t], Error [%v]",
		ERROR_NOT_REGISTED:  "parsing func is not registered: Parser tag [%s]",
		ERROR_FIELD_MATCH:   "no excel title matching found: Struct Field [%s]",
		ERROR_EINDEX_EXCEED: "the Excel column index settings exceed the line length, field [%s]",
	}

	type ErrorInfo struct {
		Row       int
		Column    string
		ErrorCode int
		ErrorMsg  string
	}
```

- `excelParser.Reader`返回的错误则是比较严重的错误，例如文件格式错误，`parser`函数未注册等等阻碍解析的严重错误。当`err != nil`时，会直接终止解析。

## 高级用法
### 自定义字段解析函数
假设你在导入文件时，你希望把Excel文件的`height`列的数据乘以2，应该如何做？
```go
type FileStruct struct {
	ID       int64     `gorm:"id"`
	Name     string    `gorm:"name" excel:"name,required"`
	Age      int8      `gorm:"age" excel:"age,required"`
	Address  string    `gorm:"address" excel:"address"`
	Birthday time.Time `gorm:"birthday" excel:"birthday,required"`
	// 1. 设置自定义解析函数标签"myheight"
	Height     float64 `gorm:"height" excel:"height,required" parser:"myheight"`
	IsStaff    bool    `gorm:"id_staff" excel:"isStaff,required"`
	Speed      int16   `gorm:"speed" excel:"speed"`
	Hobby      string  `gorm:"hobby" excel:"爱好"`
	WhatTime   int64   `gorm:"what_time" excel:"whatTime" parser:"unixNano"`
	CreateTime int64   `gorm:"create_time"`
}

func main() {
	ctx := context.Background()

	file, err := os.Open("testdata/test1.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer file.Close()

	// 2. 实现自定义解析函数
	opts := []e2s.Option{
		e2s.WithFieldParser("myheight", func(field string) (interface{}, error) {
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

	// 3. 在创建时将opts作为参数传入，自动注册
	excelParser, err := e2s.NewExcelParser("test1.xlsx", 0, "Sheet1", opts...)
	if err != nil {
		fmt.Println(err)
	}

	fileStruct := []*FileStruct{}
	// 4. 无需其他操作。可同时在opts中实现多个解析函数。
	err = excelParser.Reader(ctx, file, &fileStruct, true)
	if err != nil {
		fmt.Print(err)
	}
	for _, fs := range fileStruct {
		fmt.Printf("%+v\n", fs)
	}
}
```
得到如下输出，可见Height字段都已经成功乘以2
```shell
&{ID:0 Name:Lucas Age:18 Address:China Birthday:2005-08-17 00:00:00 +0000 UTC Height:365 IsStaff:true Speed:10 Hobby:篮球 WhatTime:1704810789000000000 CreateTime:0}
&{ID:0 Name:John Age:25 Address:USA Birthday:1999-05-07 00:00:00 +0000 UTC Height:354.16 IsStaff:false Speed:0 Hobby:足球 WhatTime:1765290790000000000 CreateTime:0}
&{ID:0 Name:Mike Age:33 Address: Birthday:2001-01-09 00:00:00 +0000 UTC Height:246 IsStaff:true Speed:0 Hobby: WhatTime:1610183951000000000 CreateTime:0}
&{ID:0 Name:Jeny Age:22 Address:Mexico Birthday:2024-02-03 00:00:00 +0000 UTC Height:912 IsStaff:false Speed:19 Hobby:羽毛球 WhatTime:0 CreateTime:0}
&{ID:0 Name:小明 Age:23 Address:中国 Birthday:2004-09-09 00:00:00 +0000 UTC Height:365 IsStaff:false Speed:12 Hobby: WhatTime:0 CreateTime:0}
```

### 标签`parser`详解
1. 当Struct的字段是以下类型时，无需额外添加`parser`
```go
		"string"
		"int"
		"int8"
		"int16"
		"int32"
		"int64"
		"float32"
		"float64"
		"bool"
		"Time": time.Time类型
```
2. 自带的`parser`函数
```go
"unixNano": 该标签函数会把时间字符串转为int64类型的纳秒时间戳，在parser标签添加即可，无需自己实现；

例如 
	WhatTime   int64     `gorm:"what_time" excel:"whatTime" parser:"unixNano"`
```

### 多线程解析
```go
opts := []e2s.Option{
		e2s.WithWorkers(12), // goroutine数量
	}
excelParser, err := e2s.NewExcelParser("test1.xlsx", 0, "Sheet1", opts...)
// 通过e2s.WithWorkers(12)设置goroutine数量为12，可实现异步解析Excel文件。
// 注意，虽然数量可以随意设置，但是真正能同时运行的goroutine数量取决于硬件的核心数，
// 所以当设置的goroutine数量超过CPU核心数时，数量再大也无意义；
```

## 导出Excel文件
当你看完以上的信息，以下的导出代码你很容易就能看懂了
```go
type Data struct {
	ID    int       `excel:"ID"`
	Name  string    `excel:"名称"`
	Value float64   `excel:"数值" convert:"mytag"`
	Date  time.Time `excel:"日期"`
}

func main() {
	ctx := context.Background()

	data := []Data{
		{ID: 1, Name: "Alice", Value: 123.45, Date: time.Now()},
		{ID: 2, Name: "Bob", Value: 67.89, Date: time.Now().AddDate(0, 0, -1)},
		{ID: 3, Name: "Charlie", Value: 99.99, Date: time.Now().AddDate(0, 0, -2)},
		{}, // 空行会输出字段类型对应的零值
	}

	opts := []e2s.WOption{
		e2s.WithFieldConverter("mytag", func(field interface{}) (interface{}, error) {
			if v, ok := field.(float64); ok {
				return 2 * math.Round(v*100) / 100, nil
			}
			return int64(0), errors.New("convert fail")
		}),
	}
	structConverter := e2s.NewStructConverter("test_write.xlsx", "./testdata/", "Sheet1", opts...)

	err := structConverter.Writer(ctx, data)
	if err != nil {
		fmt.Println(err)
	}
}
```

## 其他
1. 针对`xls`的可选参数
```go
// 在解析xls文件时，可选择文件编码
err = excelParser.Reader(ctx, file, &fileStruct, true, "utf-8") // 非必填
```
