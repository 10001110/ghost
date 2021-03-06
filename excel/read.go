package excel

import (
	"errors"
	"github.com/xuri/excelize/v2"
	"reflect"
	"strconv"
)

func Read(file, sheet string, look interface{}, skipFirstLine bool) (interface{}, error) {
	f, err := excelize.OpenFile(suffixedFileName(file))
	if err != nil {
		return nil, errors.New("open file error: " + err.Error())
	}

	rows, err := f.GetRows(sheet)
	if err != nil {
		return nil, errors.New("read data from sheet ['" + sheet + "'] error: " + err.Error())
	}

	rowOfSheet := len(rows)
	descList := descList(look)
	fileNameSet := fileNameList(descList)

	inType := reflect.TypeOf(look)

	inSliceType := reflect.SliceOf(inType)
	inTypeSlice := reflect.MakeSlice(inSliceType, 0, 0)

	var first = 1
	if skipFirstLine {
		first = 2
	}

	for i := first; i <= rowOfSheet; i++ {
		v := reflect.New(inType).Elem()
		for j, fileName := range fileNameSet {
			axis := string(rune('a'+j)) + strconv.Itoa(i)
			value, _ := f.GetCellValue(sheet, axis)
			switch (v.FieldByName(fileName).Interface()).(type) {
			case string:
				v.FieldByName(fileName).SetString(value)
			case int:
				{
					iValue, err := strconv.Atoi(value)
					if err != nil {
						continue
					}
					v.FieldByName(fileName).SetInt(int64(iValue))
				}
			case float32, float64:
				{
					fValue, err := strconv.ParseFloat(value, 64)
					if err != nil {
						continue
					}
					v.FieldByName(fileName).SetFloat(fValue)
				}
			}
		}
		inTypeSlice = reflect.Append(inTypeSlice, v)
	}
	return inTypeSlice.Interface(), nil
}
