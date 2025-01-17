package excel2struct

import (
	"context"
	"reflect"
	"sync"
)

func (ep *ExcelParser) parseWithWorkers(ctx context.Context, structFieldMetaMap map[string]FieldMetadata, titleMap map[string]int, structType, sliceType reflect.Type, outputValue reflect.Value, rows [][]string) {
	var wg sync.WaitGroup

	rowIndexChan := make(chan int, len(rows))
	resultChan := make(chan struct {
		index int
		value reflect.Value
	}, len(rows))

	// workers start
	for i := 0; i < ep.workers; i++ {
		wg.Add(1)
		SafeGo(ctx, func() {
			defer wg.Done()
			for index := range rowIndexChan {
				out := reflect.New(structType)
				parsedErr := ep.parseRowToStruct(ctx, index, structFieldMetaMap, rows[index], titleMap, out)
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

	// distribute tasks
	for i := range rows {
		rowIndexChan <- i
	}
	close(rowIndexChan)

	// wait for the processing (Map phase) to complete
	wg.Wait()
	close(resultChan)

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

func SafeGo(ctx context.Context, fn func()) {
	go func() {
		defer func() {
			// TODO catch error
			// if err := recover(); err != nil {
			// 	fmt.Println(ctx, "Recovered from panic: %v, error stack: %s", err, debug.Stack())
			// }
			_ = recover()
		}()

		fn()
	}()
}
