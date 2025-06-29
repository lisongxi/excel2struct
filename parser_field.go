package excel2struct

import (
	"errors"
	"strconv"
	"strings"
	"time"

	"github.com/shopspring/decimal"
)

var (
	DefaultFieldParserMap = map[string]FieldParser{
		"string":   FieldParserString,
		"int":      FieldParserInt,
		"int8":     FieldParserInt8,
		"int16":    FieldParserInt16,
		"int32":    FieldParserInt32,
		"int64":    FieldParserInt64,
		"float32":  FieldParserFloat32,
		"float64":  FieldParserFloat64,
		"bool":     FieldParserBool,
		"Time":     FieldParserTime,
		"unixNano": FieldParserTimeUnixNano,
	}
)

var timeLayouts = []string{
	"2006-01-02 15:04:05",
	time.RFC3339,
	"2006-01-02",
	"2006-01-02T15:04:05Z",
	"2006-01-02 15:04:05.000",
	"20060102",
	"2006/01/02",
	time.RFC1123,
	"01/02/2006",
	"01-02-06",
	"Jan 2, 2006 3:04:05 PM CST",
	time.RFC822,
	"02-Jan-2006",
	"2006-01-02T15:04:05",
	"2-Jan-06",
	"01/2/06 15:04",
	"1/2/06 15:04",
	"1/2/2006",
	"2006/1/2",
	"January 2, 2006",
	"2006-01-02 03:04:05 PM",
	"2006-01-02T15:04:05-07:00",
	"2006-01-02 15:04:05 -0700",
	"2006-01-02 15:04:05.000000",
	"2006-01-02 15:04:05.000000000",
	"20060102150405.000",
	"2006-01-02 15:04:05 MST",
	"2006-01-02 15:04:05 Z07:00",
	"1/2/2006 3:04 PM",
	"Jan 2, 2006 03:04 PM",
	"02.01.2006",
	"02.01.2006 15:04",
	"2006.01.02",
	"20060102150405",
	"150405",
	"02-Jan-06 15:04:05",
	"January 02, 2006, 03:04:05 PM",
	"Mon, 2 Jan 2006 15:04:05 -0700",
	"2006年01月02日",
	"2006年01月02日 15时04分05秒",
}

type FieldParser func(field string) (interface{}, error)

func FieldParserString(field string) (interface{}, error) {
	return field, nil
}

func FieldParserInt(field string) (interface{}, error) {
	if len(field) == 0 {
		return int(0), nil
	}
	if strings.Contains(field, ".") {
		f64, err := strconv.ParseFloat(field, 64)
		if err != nil {
			return int(0), err
		}
		return int(f64), nil
	}
	return strconv.Atoi(field)
}

func FieldParserInt8(field string) (interface{}, error) {
	if len(field) == 0 {
		return int8(0), nil
	}
	if strings.Contains(field, ".") {
		f64, err := strconv.ParseFloat(field, 64)
		if err != nil {
			return int8(0), err
		}
		return int8(f64), nil
	}
	parseInt, err := strconv.ParseInt(field, 10, 64)
	return int8(parseInt), err
}

func FieldParserInt16(field string) (interface{}, error) {
	if len(field) == 0 {
		return int16(0), nil
	}
	if strings.Contains(field, ".") {
		f64, err := strconv.ParseFloat(field, 64)
		if err != nil {
			return int16(0), err
		}
		return int16(f64), nil
	}
	parseInt, err := strconv.ParseInt(field, 10, 64)
	return int16(parseInt), err
}

func FieldParserInt32(field string) (interface{}, error) {
	if len(field) == 0 {
		return int32(0), nil
	}
	if strings.Contains(field, ".") {
		f64, err := strconv.ParseFloat(field, 64)
		if err != nil {
			return int32(0), err
		}
		return int32(f64), nil
	}
	parseInt, err := strconv.ParseInt(field, 10, 64)
	return int32(parseInt), err
}

func FieldParserInt64(field string) (interface{}, error) {
	if len(field) == 0 {
		return int64(0), nil
	}
	if strings.Contains(field, ".") {
		f64, err := strconv.ParseFloat(field, 64)
		if err != nil {
			return int64(0), err
		}
		return int64(f64), nil
	}
	return strconv.ParseInt(field, 10, 64)
}

func FieldParserFloat32(field string) (interface{}, error) {
	if len(field) == 0 {
		return float32(0.00), nil
	}
	d, err := decimal.NewFromString(field)
	if err != nil {
		return float32(0.00), err
	}
	return float32(d.Round(2).InexactFloat64()), nil
}

func FieldParserFloat64(field string) (interface{}, error) {
	if len(field) == 0 {
		return 0.00, nil
	}
	d, err := decimal.NewFromString(field)
	if err != nil {
		return 0.00, err
	}
	return d.Round(2).InexactFloat64(), nil
}

func FieldParserBool(field string) (interface{}, error) {
	if len(field) == 0 {
		return false, nil
	}
	return strconv.ParseBool(field)
}

func FieldParserTime(field string) (interface{}, error) {
	if len(field) == 0 {
		return time.Time{}, nil
	}
	for _, layout := range timeLayouts {
		t, err := time.Parse(layout, field)
		if err == nil {
			return t, nil
		}
	}
	return time.Time{}, errors.New("field time format error")
}

func FieldParserTimeUnixNano(field string) (interface{}, error) {
	if len(field) == 0 {
		return int64(0), nil
	}
	for _, layout := range timeLayouts {
		t, err := time.Parse(layout, field)
		if err == nil {
			return t.UnixNano(), nil
		}
	}
	return int64(0), errors.New("field time UNIX Nano format error")
}
