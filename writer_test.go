package excel2struct

import (
	"context"
	"errors"
	"math"
	"testing"
	"time"

	"github.com/zeebo/assert"
)

type Data struct {
	ID    int       `excel:"ID"`
	Name  string    `excel:"名称"`
	Value float64   `excel:"数值" convert:"mytag"`
	Date  time.Time `excel:"日期"`
}

func TestWriter(t *testing.T) {
	ctx := context.Background()

	data := []Data{
		{ID: 1, Name: "Alice", Value: 123.45, Date: time.Now()},
		{ID: 2, Name: "Bob", Value: 67.89, Date: time.Now().AddDate(0, 0, -1)},
		{ID: 3, Name: "Charlie", Value: 99.99, Date: time.Now().AddDate(0, 0, -2)},
		{},
	}

	opts := []WOption{
		WithFieldConverter("mytag", func(field interface{}) (interface{}, error) {
			if v, ok := field.(float64); ok {
				return 2 * math.Round(v*100) / 100, nil
			}
			return int64(0), errors.New("convert fail")
		}),
	}
	structConverter := NewStructConverter("test_write.xlsx", "./testdata/", "Sheet1", opts...)

	err := structConverter.Writer(ctx, data)

	assert.Nil(t, err)
}
