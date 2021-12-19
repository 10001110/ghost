package excel

import (
	"reflect"
	"sort"
	"strconv"
	"strings"
)

type Desc struct {
	O         int    `json:"o"`
	W         int    `json:"w"`
	T         string `json:"t"`
	FiledName string
}

func descList(in interface{}) []Desc {
	t := reflect.TypeOf(in)
	var ret = make([]Desc, 0)
	for i := 0; i < t.NumField(); i++ {
		f := t.Field(i)
		xlsxTag := f.Tag.Get("xlsx")
		if xlsxTag == "" {
			continue
		}

		desc := fromString(xlsxTag)
		desc.FiledName = f.Name
		ret = append(ret, desc)
	}
	sort.SliceStable(ret, func(i, j int) bool {
		return ret[i].O < ret[j].O
	})
	return ret
}

func titleList(descList []Desc) []interface{} {
	var ret = make([]interface{}, 0)
	for _, v := range descList {
		if v.T == "" {
			// if there is no title defined use fileName
			ret = append(ret, v.FiledName)
		} else {
			ret = append(ret, v.T)
		}
	}
	// fmt.Println("title:", ret)
	return ret
}

func fileNameList(descList []Desc) []string {
	var ret = make([]string, 0)
	for _, v := range descList {
		ret = append(ret, v.FiledName)
	}
	return ret
}

func widthList(descList []Desc) []int {

	var ret = make([]int, 0)

	// if w != 0, use it
	// if w == 0, use fileName width
	for _, v := range descList {
		if v.W != 0 {
			ret = append(ret, v.W)
		} else {
			ret = append(ret, len(v.FiledName))
		}
	}

	return ret
}

func fromString(desc string) Desc {
	var t Desc
	split := strings.Split(desc, ";")
	for _, v := range split {
		i := strings.Split(v, ":")
		switch i[0] {
		case "o":
			{
				o, _ := strconv.Atoi(i[1])
				t.O = o
			}
		case "w":
			{
				w, _ := strconv.Atoi(i[1])
				t.W = w
			}
		case "t":
			t.T = i[1]
		}
	}
	return t
}

func suffixedFileName(fileName string) string {
	if strings.HasSuffix(fileName, "xlsx") || strings.HasSuffix(fileName, "xls") {
		return fileName
	}
	return fileName + ".xlsx"
}
