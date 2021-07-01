package xlsxToMail

import (
	"bytes"
	"database/sql"
	"encoding/base64"
	"fmt"
    "github.com/tealeg/xlsx"
	"net/smtp"
	"strings"
	"time"

)

type MailServer struct {
	Smtp     string
	Port     string
	Username string
	Password string
}

type MailSetting struct {
	Subject string
	From    string
	To      []string
	Charset string
}

type ExelFileSetting struct {
	SheetName string
}

type Xtm struct {
	MailServer       MailServer
	MailSettings     MailSetting
	ExelFileSettings ExelFileSetting
}

func New() *Xtm {
	return &Xtm{}
}

func (xtm *Xtm) generateExel(rows sql.Rows) error {
	colNames, err := rows.Columns()
	defer rows.Close()
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
	xsheet, err := xfile.AddSheet(xtm.ExelFileSettings.SheetName)
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
	attach := base64.StdEncoding.EncodeToString(b.Bytes())
	err = xtm.sendMail(attach)
	if err != nil {
		return err

	}

	return nil
}

func (xtm *Xtm) sendMail(attach string) error {


	To := strings.Join(xtm.MailSettings.To, ",")
	Subject := base64.StdEncoding.EncodeToString([]byte(xtm.MailSettings.Subject))
	message := "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\nContent-Disposition: attachment; charset=" + xtm.MailSettings.Charset + ";filename=reports.xlsx\nContent-Transfer-Encoding: base64\nFrom: " + xtm.MailSettings.From + "\n" + "Subject: =?" + xtm.MailSettings.Charset + "?B?" + Subject + "?=\n" + "To: " + To + "\n\n" + attach
	auth := smtp.PlainAuth("", xtm.MailSettings.From, xtm.MailServer.Password, xtm.MailServer.Smtp)
    err:= smtp.SendMail(xtm.MailServer.Smtp+":"+xtm.MailServer.Port, auth, xtm.MailSettings.From, xtm.MailSettings.To, []byte(message))
	if err != nil{
		return err
	}
	return nil

}

func (xtm *Xtm) SendFile(rows sql.Rows) {

	err := xtm.generateExel(rows)
	if err != nil {
		fmt.Println(err)
	}

}
