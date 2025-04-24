package excel2struct

import (
	"context"
	"reflect"
	"sync"

	"github.com/lisongxi/goutils"
)

func (ep *ExcelParser) parseWithWorkers(ctx context.Context, structFieldMetaMap map[string]FieldMetadata, titleMap map[string]int, structType, sliceType reflect.Type, outputValue reflect.Value, rows [][]string, skip bool) {
	defer close(ep.errChan)
	var wg sync.WaitGroup

	rowIndexChan := make(chan int, ep.workers)
	resultChan := make(chan struct {
		index int
		value reflect.Value
	}, ep.workers)

	// workers start
	for i := 0; i < ep.workers; i++ {
		wg.Add(1)
		goutils.SafeGo(ctx, func() {
			defer wg.Done()
			for index := range rowIndexChan {
				out := reflect.New(structType)
				parsedErr := ep.parseRowToStruct(ctx, index+ep.headerIndex+2, structFieldMetaMap, rows[index], titleMap, out, skip)
				if parsedErr != nil {
					continue
				}
				resultChan <- struct {
					index int
					value reflect.Value
				}{index: index, value: out}
			}
		})
	}

	// 分发任务 distribute tasks
	go func() {
		for i := range rows {
			rowIndexChan <- i
		}
		close(rowIndexChan)
	}()

	// 等待任务完成 wait for task to complete
	go func() {
		wg.Wait()
		close(resultChan) // 只要resultChan未关闭，下面的range resultChan就一直循环读取
	}()

	// Reduce phase: Collect results
	resultMap := make(map[int]reflect.Value)
	for res := range resultChan {
		resultMap[res.index] = res.value
	}

	// End Result
	results := reflect.MakeSlice(sliceType, 0, len(rows))
	for i := 0; i < len(rows); i++ {
		if value, ok := resultMap[i]; ok {
			results = reflect.Append(results, value)
		}
	}
	outputValue.Elem().Set(results)
}
