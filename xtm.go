package xlsxToMail

import (
	"bytes"
	"database/sql"
	"fmt"
	"github.com/tealeg/xlsx"
	"time"
)

type MailServer struct {
	Smtp     string
	Port     string
	Username string
	Password string
}

type Mail struct {
	Subject string
	From    string
	To      []string
	Charset string
}

type ExelFileSetting struct {
	Sheet *xlsx.Sheet
}

type Xtm struct {
	MailServer
	Mail
	ExelFileSetting
}

func (xtm *Xtm) Init() *Xtm {

	return xtm
}

func (xtm *Xtm) GenerateExel(rows sql.Rows) error {

	colNames, err := rows.Columns()
	if err != nil {
		return fmt.Errorf("error fetching column names, %s\n", err)
	}
	length := len(colNames)
	pointers := make([]interface{}, length)
	container := make([]interface{}, length)
	for i := range pointers {
		pointers[i] = &container[i]
	}
	xfile := xlsx.NewFile()
	xsheet, err := xfile.AddSheet(xtm.ExelFileSetting.Sheet.Name)
	if err != nil {
		return fmt.Errorf("error adding sheet to xlsx file, %s\n", err)
	}
	xsheet.SetColWidth(1, 4, 20.0)
	xrow := xsheet.AddRow()
	xrow.WriteSlice(&colNames, -1)
	for rows.Next() {
		err = rows.Scan(pointers...)
		if err != nil {
			return fmt.Errorf("error scanning sql row, %s\n", err)
		}
		xrow = xsheet.AddRow()
		for _, v := range container {
			xcell := xrow.AddCell()
			switch v := v.(type) {
			case string:
				xcell.SetString(v)
			case []byte:
				xcell.SetString(string(v))
			case int64:
				xcell.SetInt64(v)
			case float64:
				xcell.SetFloat(v)
			case bool:
				xcell.SetBool(v)
			case time.Time:
				xcell.SetDateTime(v)
			default:
				xcell.SetValue(v)
			}

		}

	}
	var b bytes.Buffer
	xfile.Write(&b)
	//attach := base64.StdEncoding.EncodeToString(b.Bytes())
	//err = conf.sendMail(attach)
	//if err != nil {
	//	logger("fatal", err)

	//}
	//return nil
	return nil
}

func (xtm *Xtm) SendExel(rows sql.Rows) {

	err := xtm.GenerateExel(rows)
	if err != nil {
		fmt.Println(err)
	}

}
