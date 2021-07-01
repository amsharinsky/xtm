# xlsxToMail

        sqlQuery := `SELECT * FROM table`
	rows,_:=conn.Query(sqlQuery)
	mail:=xlsxToMail.New()
	mail.MailServer=xlsxToMail.MailServer{
		Smtp: "smtp.yandex.ru",
		Port: "465",
		Username: "login@yandex.ru",
		Password:"dsfsdfsfsd",

	}

	mail.MailSettings=xlsxToMail.MailSetting{
		Subject: "test",
		From:    "login@yandex.ru",
		To:      []string{"dd@mail.ru"},
		Charset: "utf-8",
	}
	mail.ExelFileSettings=xlsxToMail.ExelFileSetting{
		SheetName: "test",
	}
    mail.SendFile(*rows)
