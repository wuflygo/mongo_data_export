package main

import (
	"bufio"
	"flag"
	"fmt"
	"github.com/duke-git/lancet/v2/convertor"
	list "github.com/duke-git/lancet/v2/datastructure/list"
	"github.com/duke-git/lancet/v2/datetime"
	jsoniter "github.com/json-iterator/go"
	"github.com/xuri/excelize/v2"
	"io"
	"os"
	"path"
	"regexp"
	"strings"
)

var (
	regexpList = map[string]string{`ISODate\(`: "", `ObjectId\(`: "", `\"\)`: `"`}
)

func main() {
	var filePath = flag.String("c", "", "config path for flag c")
	flag.Parse()
	fmt.Println("ip has value ", *filePath)
	if filePath == nil || *filePath == "" {
		panic("文件不能为空")
	}
	lines, err := ReadLine(*filePath)
	if err != nil {
		panic(fmt.Sprintf("打开文件失败：%s", err))
	}

	headers := list.NewList([]string{})

	lineMapList := make([]map[string]interface{}, 0)

	var json = jsoniter.ConfigCompatibleWithStandardLibrary
	for _, line := range lines {
		lineJson := filterLine(line)
		var lineMap map[string]interface{}
		err := json.Unmarshal([]byte(lineJson), &lineMap)
		if err != nil {
			panic(fmt.Sprintf("数据转换出错：%s", err))
		}
		for header, _ := range lineMap {
			if !headers.Contain(header) {
				headers.Push(header)
			}
		}

		for k, v := range lineMap {
			switch v.(type) {
			case map[string]interface{}:
				nv := v.(map[string]interface{})
				if rv, ok := nv["$oid"]; ok {
					lineMap[k] = rv
				}
				if rv, ok := nv["$date"]; ok {
					lineMap[k] = datetime.NewUnix(int64(int(rv.(float64)) / 1000)).ToFormat()
				}

			}
		}

		lineMapList = append(lineMapList, lineMap)
	}

	// 创建填充文件
	f := excelize.NewFile()
	// 创建一个工作表
	index := f.NewSheet("Sheet1")
	// 设置单元格的值
	for i, headerName := range headers.Data() {
		f.SetCellValue("Sheet1", NumTransferStr(i+1)+"1", headerName)
	}

	for i, line := range lineMapList {
		lineNum := i + 2
		for j, headerName := range headers.Data() {
			if v, ok := line[headerName]; ok {
				f.SetCellValue("Sheet1", NumTransferStr(j+1)+convertor.ToString(lineNum), v)
			}
		}
	}

	// 设置工作簿的默认工作表
	f.SetActiveSheet(index)
	// 根据指定路径保存文件
	//获取文件名称带后缀
	fileNameWithSuffix := path.Base(*filePath)
	//获取文件的后缀(文件类型)
	fileType := path.Ext(fileNameWithSuffix)
	//获取文件名称(不带后缀)
	fileNameOnly := strings.TrimSuffix(fileNameWithSuffix, fileType)
	if err := f.SaveAs(fileNameOnly + ".xlsx"); err != nil {
		fmt.Println(err)
	}
}

func filterLine(line string) string {
	for rex, repl := range regexpList {
		sampleRegexp := regexp.MustCompile(rex)
		line = sampleRegexp.ReplaceAllString(line, repl)
	}
	return line
}

func ReadLine(fileName string) ([]string, error) {
	f, err := os.Open(fileName)
	if err != nil {
		return nil, err
	}
	buf := bufio.NewReader(f)
	var result []string
	for {
		line, err := buf.ReadString('\n')
		line = strings.TrimSpace(line)
		result = append(result, line)
		if err != nil {
			if err == io.EOF { //读取结束，会报EOF
				return result, nil
			}
			return nil, err
		}
	}
	return result, nil
}

func NumTransferStr(Num int) string {
	var (
		Str  string = ""
		k    int
		temp []int //保存转化后每一位数据的值，然后通过索引的方式匹配A-Z
	)
	//用来匹配的字符A-Z
	Slice := []string{"", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O",
		"P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

	if Num > 26 { //数据大于26需要进行拆分
		for {
			k = Num % 26 //从个位开始拆分，如果求余为0，说明末尾为26，也就是Z，如果是转化为26进制数，则末尾是可以为0的，这里必须为A-Z中的一个
			if k == 0 {
				temp = append(temp, 26)
				k = 26
			} else {
				temp = append(temp, k)
			}
			Num = (Num - k) / 26 //减去Num最后一位数的值，因为已经记录在temp中
			if Num <= 26 {       //小于等于26直接进行匹配，不需要进行数据拆分
				temp = append(temp, Num)
				break
			}
		}
	} else {
		return Slice[Num]
	}

	for _, value := range temp {
		Str = Slice[value] + Str //因为数据切分后存储顺序是反的，所以Str要放在后面
	}
	return Str
}
