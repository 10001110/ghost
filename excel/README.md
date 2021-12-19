```golang
package main

import (
	"encoding/json"
	"fmt"
	"github.com/10001110/ghost/excel"
)

// SR
// o -> order
// w -> width
// t -> title
type SR struct {
	EventTime        string  `xlsx:"o:4;w:10;t:时间"`
	SearchContent    string  `xlsx:"o:1;w:20;t:内容"`
	SearchResultsYes string  `xlsx:"o:2;w:20;t:是否有搜索结果"`
	AreaCode         int     `xlsx:"o:3;w:20;t:地区码"`
	ASD              float64 `xlsx:"o:5;w:20;t:asd"`
	Ignore           string  // 这个字段没有xlsx标签，不会写入到excel,
}

func main() {
	var s = make([]SR, 0)

	sr01 := SR{"2021-12-08", "hello", "no", 110110, 12.4, ""}
	sr02 := SR{"2021-12-08", "shutdown -h now", "no", 119119, 3.14, ""}
	sr03 := SR{"2021-12-08", "teacher guo", "yes", 112114, 2.717, "this will be ignored"}
	sr04 := SR{"2021-12-08", "oh oh oh", "no", 911911, 69.411, ""}
	s = append(s, sr01, sr02, sr03, sr04)

	// create excel
	where := excel.Create("/Users/haonan/go/src/tmp/t", "search content", SR{}, s)
	println(where)

	// read excel
	slice, _ := excel.Read("/Users/haonan/go/src/tmp/t", "search content", SR{}, true)
	for _, v := range slice.([]SR) {
		marshal, _ := json.Marshal(v)
		fmt.Println(string(marshal))
	}
}
```