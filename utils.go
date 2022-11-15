package main

import (
	"fmt"
	"strconv"

	"github.com/araddon/dateparse"
	"github.com/xuri/excelize/v2"
)


func changeDate(date string) string {
	// var error bool = false
	var year string
	var day string
	var month string
	var dayB bool = true
	var monthB bool = true
	for _, char := range date {
		if char != '/' && dayB {
			day += string(char)
		} else if char == '/' && dayB {
			dayB = false
		} else if char != '/' && monthB {
			month += string(char)
		} else if char == '/' && monthB {
			monthB = false
		} else if !monthB && !dayB {
			year += string(char)
		}
	}
	return month + "/" + day + "/" + year
}

func toChar(i int) rune {
	return rune('A' - 1 + i)
}

func SaveToExcel(data [][]string) error {
	f := excelize.NewFile()
	dict := map[string]string{
		"A1": "N п/п",
		"B1": "Дата",
		"C1": "Время",
		"D1": "Работник 1",
		"E1": "Работник 2",
		"F1": "Работник 3",
		"G1": "Тип объекта",
		"H1": "Тип уборки",
		"I1": "Доп. услуги",
		"J1": "Расходн.",
		"K1": "Оборудование",
		"L1": "Цена",
		"M1": "Стоимость",
		"N1": "Счет",
		"O1": "Ф-ма опл.",
		"P1": "Работникам",
		"Q1": "Заметки",
	}
	for key := range dict {
		f.SetCellValue("Sheet1", key, dict[key])
	}

	for i := 0; i < len(data); i++ {
		for j := 0; j < len(data[i]); j++ {
			char := toChar(j + 1)
			f.SetCellValue("Sheet1", string(char)+strconv.Itoa(i+2), data[i][j])
		}

	}

	if err := f.SaveAs(cnf.saveEntry.Text); err != nil {
		return err
	}
	cnf.startButton.Enable()
	cnf.loadButton.Enable()
	cnf.textDone.Hidden = false
	cnf.progressBar.SetValue(1)
	return nil

}

func sortByDate(data [][] string) [][] string{
	for {
		flag := true
		for i := 0; i < len(data); i++ {
			for j := 0; j < len(data[i]); j++ {
				if j == 1 {
					if len(data) != i+1 {
						fDate ,_ :=dateparse.ParseAny(changeDate(data[i][j]))
						sDate ,_ :=dateparse.ParseAny(changeDate(data[i+1][j]))
						if fDate.Unix() < sDate.Unix(){
							data[i], data[i+1] = data[i+1], data[i]
							flag = false
						}
					}
				}
			}
		}
		if flag {
			break
		}
	}
return data
}

func (app *config) getData(){
	go func() {
		app.startButton.Disable()
		app.loadButton.Disable()
		app.textDone.Hidden = true
		var dataSlice [][]string
		startDate, _ := dateparse.ParseAny(app.startDate.Text)
		endDate, _ := dateparse.ParseAny(app.endDate.Text)
		

		file, err := excelize.OpenFile(app.loadEntry.Text)
		if err != nil {
			fmt.Println(err)
			return
		}

		defer func() {
			// Close the spreadsheet.
			if err := file.Close(); err != nil {
				fmt.Println(err)
			}
		}()
		// Get sheets list from a excel file.
		sheetlist := file.GetSheetList()
		// app.sheetList.SetText("Sheets: " + "0" + "/" + strconv.Itoa(len(sheetlist)))
		
			for in, sheet := range sheetlist {
				app.progressBar.SetValue(float64(in) / float64(len(sheetlist)))
				app.sheetList.SetText("Sheets: " + strconv.Itoa(in+1) + "/" + strconv.Itoa(len(sheetlist)))
				cols, err := file.GetCols(sheet)
				if err != nil {
					fmt.Println(err)
					return
				}
				
				for i, col := range cols {
					for _, colCell := range col {
						date := "Дата"
						if date == colCell {
							for j, colCell := range cols[i] {
								_, err := dateparse.ParseAny(colCell)
								if err != nil {
									continue
								}
								colCell = changeDate(colCell)
								t, _ := dateparse.ParseAny(colCell)
								timestamp := t.Unix()
								if timestamp >= startDate.Unix() && timestamp <= endDate.Unix() {
									rows, err := file.GetRows(sheet)
									if err != nil {
										fmt.Println(err)
										return
									}
									dataSlice = append(dataSlice, rows[j])
		
								}
							}
						}
				}
	
			}
		
		}
	sortedData := sortByDate(dataSlice)
	SaveToExcel(sortedData)
	}()
}

func (app *config) loadSheets(){
	file, err := excelize.OpenFile(app.loadEntry.Text)
	if err != nil {
		fmt.Println(err)
		return
	}

	defer func() {
		// Close the spreadsheet.
		if err := file.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// Get sheets list from a excel file.
	sheetlist := file.GetSheetList()
	app.sheetList.SetText("Sheets: " + "0" + "/" + strconv.Itoa(len(sheetlist)))
}