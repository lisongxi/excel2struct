package excel2struct

import (
	"errors"
	"strconv"
	"strings"
	"time"
)

var (
	DefaultFieldParserMap = map[string]FieldParser{
		"string":  FieldParserString,
		"int":     FieldParserInt,
		"int8":    FieldParserInt8,
		"int16":   FieldParserInt16,
		"int32":   FieldParserInt32,
		"int64":   FieldParserInt64,
		"float32": FieldParserFloat32,
		"float64": FieldParserFloat64,
		"bool":    FieldParserBool,
		"Time":    FieldParserTime,
		"timeUN":  FieldParserTimeUnixNano,
	}

	TimeLayoutParserMap = map[string]TimeLayoutParser{
		"time_lay":    FieldParserTimeWithLayout,
		"time_lay_un": FieldParserTimeWithLayoutUnixNano,
	}

	TimeLocLayoutParserMap = map[string]TimeLocLayoutParser{
		"time_loc_lay":    FieldParserTimeWithLayoutLoc,
		"time_loc_lay_un": FieldParserTimeWithLayoutLocUnixNano,
	}
)

var timeLayouts = []string{
	"2006-01-02 15:04:05",
	"2006-01-02",
	"2006-01-02 15:04:05.000",
	"2006/01/02",
	"01/02/2006",
	"01-02-06",
	"2-Jan-06",
	"2006-01-02T15:04:05Z",
	"2006-01-02T15:04:05",
	"Jan 2, 2006 3:04:05 PM CST",
	"2006-01-02 15:04:05 -0700",
	"02-Jan-2006",
	"01/2/06 15:04",
	"1/2/2006",
	"2006/1/2",
	"January 2, 2006",
}

type FieldParser func(field string) (interface{}, error)

type TimeLayoutParser func(field, layout string) (interface{}, error)

type TimeLocLayoutParser func(field string, loc *time.Location, layout string) (interface{}, error)

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
	f64, err := strconv.ParseFloat(field, 32)
	if err != nil {
		return float32(0.00), err
	}
	return float32(f64), nil
}

func FieldParserFloat64(field string) (interface{}, error) {
	if len(field) == 0 {
		return 0.00, nil
	}
	return strconv.ParseFloat(field, 32)
}

func FieldParserBool(field string) (interface{}, error) {
	if len(field) == 0 {
		return false, nil
	}
	return strconv.ParseBool(field)
}

func FieldParserTime(field string) (interface{}, error) {
	if len(field) == 0 {
		return "", nil
	}
	for _, layout := range timeLayouts {
		t, err := time.Parse(layout, field)
		if err == nil {
			return t, nil
		}
	}
	return time.Time{}, errors.New("field time format error")
}

func FieldParserTimeWithLayout(field string, layout string) (interface{}, error) {
	if len(field) == 0 {
		return "", nil
	}
	t, err := time.Parse(layout, field)
	if err != nil {
		return "", err
	}
	return t.Format(layout), nil
}

func FieldParserTimeWithLayoutLoc(field string, loc *time.Location, layout string) (interface{}, error) {
	if len(field) == 0 {
		return "", nil
	}
	t, err := time.ParseInLocation(layout, field, loc)
	if err != nil {
		return "", err
	}
	return t.Format(layout), nil
}

func FieldParserTimeUnixNano(field string) (interface{}, error) {
	if len(field) == 0 {
		return 0, nil
	}
	t, err := time.Parse("2006-01-02 15:04:05", field)
	if err != nil {
		return 0, err
	}
	return t.UnixNano(), nil
}

func FieldParserTimeWithLayoutUnixNano(field string, layout string) (interface{}, error) {
	if len(field) == 0 {
		return 0, nil
	}
	t, err := time.Parse(layout, field)
	if err != nil {
		return 0, err
	}
	return t.UnixNano(), nil
}

func FieldParserTimeWithLayoutLocUnixNano(field string, loc *time.Location, layout string) (interface{}, error) {
	if len(field) == 0 {
		return 0, nil
	}
	t, err := time.ParseInLocation(layout, field, loc)
	if err != nil {
		return 0, err
	}
	return t.UnixNano(), nil
}
