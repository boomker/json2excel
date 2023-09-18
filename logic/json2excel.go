package logic

import (
	"encoding/json"
	"fmt"
	"json2excel/common"
	"json2excel/log"
	"time"

	// "github.com/360EntSecGroup-Skylar/excelize/v2"
	orderedmap "github.com/wk8/go-ordered-map/v2"
	"github.com/xuri/excelize/v2"
)

// github.com/360EntSecGroup-Skylar/excelize/v2
// document: https://xuri.me/excelize/zh-hans/

type Json2Excel struct{}

func (j *Json2Excel) Json2Excel(jsonBytes []byte, saveDir string) (savePath string, err error) {
	jsonMap := make([]map[string]interface{}, 0)
	err = json.Unmarshal(jsonBytes, &jsonMap)
	if err != nil {
		log.Errorf("json.Unmarshal error. err: %s", err.Error())
		return "", err
	}

	var jrm []json.RawMessage
	_ = json.Unmarshal(jsonBytes, &jrm)

	headers := []string{}
	om := orderedmap.New[string, any]()
	for _, r := range jrm {
		if len(r) < 1 {
			log.Error("json data count lt 1")
			return "", fmt.Errorf("json data count lt 1")
		}

		err := json.Unmarshal([]byte(r), &om)
		if err != nil {
			log.Errorf("json.Unmarshal error. err: %s", err.Error())
			return "", err
		}
	}
	for pair := om.Oldest(); pair != nil; pair = pair.Next() {
		headers = append(headers, pair.Key)
	}

	// header := make([]string, 10)
	// // temp := map[string]struct{}{}
	// idx := 0
	// for i, v := range jsonMap {
	// 	for k := range v {
	// 		println(k)
	// 		header[idx] = k
	// 		idx++
	// 		// if _, ok := temp[km]; !ok {
	// 		// 	temp[km] = struct{}{}
	// 		// 	header = append(header, km)
	// 		// }
	// 	}
	// 	if i == 0 {
	// 		break
	// 	}
	// }

	// sort.Strings(header)

	f := excelize.NewFile()

	sIndexX := "B"
	eIndexX, err := excelize.ColumnNumberToName(int(sIndexX[0]) + len(headers) - int('A'))
	if err != nil {
		log.Errorf("excelize.ColumnNumberToName error. err: %s", err.Error())
		return "", err
	}
	indexY := 2
	sheetName := "Sheet1"

	//设置列宽
	if err = f.SetColWidth(sheetName, sIndexX, eIndexX, 20); err != nil {
		log.Errorf("f.SetColWidth error. err: %s", err.Error())
		return "", err
	}

	//设置header行高
	if err = f.SetRowHeight(sheetName, indexY, 20); err != nil {
		log.Errorf("f.SetRowHeight error. err: %s", err.Error())
		return "", err
	}

	//设置头
	// headerStyle, err := f.NewStyle(`{
	// 	"font":
	// 	{
	// 		"bold": true,
	// 		"size": 14
	// 	},
	// 	"alignment":
	// 	{
	// 		"horizontal": "left",
	// 		"vertical": "center"
	// 	}
	// }`)

	headerStyle, err := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Color: "1f7f3b", Bold: true, Family: "Microsoft YaHei"},
		Alignment: &excelize.Alignment{Vertical: "center", Horizontal: "center"},
		Border:    []excelize.Border{{Type: "top", Style: 2, Color: "1f7f3b"}},
	})
	if err != nil {
		return "", err
	}

	err = f.SetCellStyle(sheetName, fmt.Sprintf("%s%d", sIndexX, indexY), fmt.Sprintf("%s%d", eIndexX, indexY), headerStyle)
	if err != nil {
		log.Errorf("f.SetCellStyle error. err: %s", err.Error())
		return "", err
	}

	err = f.SetSheetRow(sheetName, fmt.Sprintf("%s%d", sIndexX, indexY), &headers)
	if err != nil {
		log.Errorf("f.SetSheetRow error. err: %s", err.Error())
		return "", err
	}
	indexY++

	//设置值
	values := make([]interface{}, len(headers))
	for _, row := range jsonMap {
		for i, vm := range headers {
			if val, ok := row[vm]; ok {
				values[i] = common.TransCellVal(val)
			} else {
				values[i] = nil
			}
		}

		err = f.SetSheetRow(sheetName, fmt.Sprintf("%s%d", sIndexX, indexY), &values)
		if err != nil {
			log.Errorf("f.SetSheetRow error. err: %s", err.Error())
			return "", err
		}
		indexY++
	}

	savePath = fmt.Sprintf("%s/%s.xlsx", saveDir, time.Now().Format("20060102150405"))
	if err := f.SaveAs(savePath); err != nil {
		log.Errorf("f.SaveAs error. savePath: %s, err: %s", savePath, err.Error())
		return "", err
	}

	return savePath, nil
}
