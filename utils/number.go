package utils

import (
	"regexp"
)

func IsNumber(s string) bool {
	matched, _ := regexp.MatchString(`^-?\d+$`, s)
	if matched {
		return true
	}

	matched, _ = regexp.MatchString(`^-?\d+(\.\d+)?$`, s)
	return matched
}
