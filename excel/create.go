package excel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"reflect"
	"strconv"
)

func Create(filename, sheetName string, look, inputDataSet interface{}) string {
	filename = suffixedFileName(filename)
	tmpFile := excelize.NewFile()

	sheet := tmpFile.NewSheet(sheetName)
	tmpFile.DeleteSheet("Sheet1")

	descList := descList(look)
	titleList := titleList(descList)
	widthList := widthList(descList)

	switch reflect.TypeOf(inputDataSet).Kind() {
	case reflect.Slice:
		dataSet := reflect.ValueOf(inputDataSet)
		tmpFile.SetSheetRow(sheetName, "A1", &titleList)
		fileNameSet := fileNameList(descList)
		tmpFile.InsertRow(sheetName, dataSet.Len()+1)
		for i := 0; i < dataSet.Len(); i++ {
			var rowStuffing = make([]interface{}, 0)
			for _, curFiledName := range fileNameSet {
				rowStuffing = append(rowStuffing, dataSet.Index(i).FieldByName(curFiledName))
			}
			tmpFile.SetSheetRow(sheetName, "A"+strconv.Itoa(i+2), &rowStuffing)
			println()
		}
	}

	for i := 0; i < len(widthList); i++ {
		tmpFile.SetColWidth(sheetName, string(rune('A'+i)), string(rune('A'+i+1)), float64(widthList[i]))
	}

	tmpFile.SetActiveSheet(sheet)
	if err := tmpFile.SaveAs(filename); err != nil {
		fmt.Println(err)
	}

	return filename
}
