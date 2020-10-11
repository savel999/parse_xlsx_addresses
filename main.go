package main

import (
	"fmt"
	"regexp"
	"strings"
	"time"

	"github.com/cheggaaa/pb/v3"
	"github.com/tealeg/xlsx"
	"gopkg.in/webdeskltd/dadata.v2"
)

var ColumnsPriority = []int{3, 2}
var AddressPrefix = "Новосибирск, "
var DadataToken = "53c59eb868e64e84d220fd718853f2789e83b479"

func processAddress(adr string) string {
	var result string

	var re = regexp.MustCompile(`(?i)((.+)\(|(.+))`)
	matches := re.FindStringSubmatch(adr)
	matchesCount := len(matches)

	if matchesCount == 4 {
		if matches[2] != "" {
			result = AddressPrefix + matches[2]
		} else if matches[3] != "" {
			result = AddressPrefix + matches[3]
		}
	}

	result = strings.TrimSpace(result)

	return result
}

func main() {
	wb, err := xlsx.OpenFile("input.xlsx")

	if err != nil {
		panic(err)
	}

	sh := wb.Sheets[0]

	currIndex := 0

	//клиент дадаты
	dadataClient := dadata.NewDaData(DadataToken, "")

	//стиль корректной ячейки
	correctStyle := xlsx.NewStyle()
	correctStyle.Fill.FgColor = "FFC6EFCE"
	correctStyle.Fill.PatternType = "solid"

	//стиль некорректной ячейки
	incorrectStyle := xlsx.NewStyle()
	incorrectStyle.Fill.FgColor = "FFFFC7CE"
	incorrectStyle.Fill.PatternType = "solid"

	bar := pb.Full.Start(sh.MaxRow)

	for rowIndex := 0; rowIndex < sh.MaxRow; rowIndex++ {
		bar.Increment()

		if currIndex > 5 {
			break
		}

		currentRow, err := sh.Row(rowIndex)

		if err != nil {
			continue
		}

		if rowIndex == 0 {
			newCell := currentRow.AddCell()
			newCell.Value = "Адрес DaData"

			newCell2 := currentRow.AddCell()
			newCell2.Value = "Уровень фиас"

			newCell3 := currentRow.AddCell()
			newCell3.Value = "Проверка DaData"
			continue
		}

		isCorrectAddress := false

		for _, priority := range ColumnsPriority {
			cellValue := currentRow.GetCell(priority - 1).Value

			processedValue := processAddress(cellValue)

			if processedValue != "" {

				//проверка Адреса
				params := dadata.SuggestRequestParams{Query: processedValue, Count: 1}
				resultDadata, errDadata := dadataClient.SuggestAddresses(params)

				if errDadata != nil {
					fmt.Println(errDadata)
				} else {
					if len(resultDadata) > 0 {
						firstItem := resultDadata[0]
						curCellFiasVal := ""

						if firstItem.Data.HouseFiasID != "" {
							curCellFiasVal = firstItem.Data.HouseFiasID
						} else if firstItem.Data.FiasID != "" {
							curCellFiasVal = firstItem.Data.FiasID
						}

						newCell := currentRow.AddCell()
						newCell.Value = firstItem.Value

						newCell2 := currentRow.AddCell()
						newCell2.Value = firstItem.Data.FiasLevel

						if firstItem.Data.FiasLevel == "8" {
							newCell.SetStyle(correctStyle)
						} else {
							newCell.SetStyle(incorrectStyle)
						}

						newCell3 := currentRow.AddCell()
						newCell3.Value = curCellFiasVal
						isCorrectAddress = true
						break
					}

				}
			}
		}

		if !isCorrectAddress {
			newCell := currentRow.AddCell()
			newCell.Value = "Не найден"
			newCell.SetStyle(incorrectStyle)
		}

		//currIndex++

		time.Sleep(time.Millisecond * 300)
	}

	bar.Finish()

	fmt.Println("Конец")

	err = wb.Save("output.xlsx")

	if err != nil {
		fmt.Println(err)
	}
}
