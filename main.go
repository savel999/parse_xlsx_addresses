package main

import (
	"fmt"
	"github.com/cheggaaa/pb/v3"
	"github.com/tealeg/xlsx"
	"gopkg.in/webdeskltd/dadata.v2"
	"regexp"
	"strings"
	"time"
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
			newCell.Value = "Проверка DaData"
			continue
		}
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
						newCell.Value = curCellFiasVal
						break
					}

				}
			}
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
