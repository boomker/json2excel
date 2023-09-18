package common

import (
	"encoding/json"
	"os"
	"reflect"
	"strconv"
)

func ReadFile(filePath string) ([]byte, error) {
	// f, err := os.Open(filePath)
	f, err := os.ReadFile(filePath)
	if err != nil {
		return []byte{}, err
	}

	return f, nil
}

func TransCellVal(val interface{}) (v interface{}) {
	valRef := reflect.ValueOf(val)
	if !valRef.IsValid() {
		v = nil
	} else if valRef.Kind() == reflect.Slice || valRef.Kind() == reflect.Array || valRef.Kind() == reflect.Map || valRef.Kind() == reflect.Struct || valRef.Kind() == reflect.Ptr {
		valjs, _ := json.Marshal(val)
		v = string(valjs)
	} else if valRef.Kind() == reflect.Bool {
		v = strconv.FormatBool(val.(bool))
	} else {
		v = val
	}

	return v
}
