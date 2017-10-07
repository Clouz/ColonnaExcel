package main

import (
	"fmt"
	"reflect"

	"github.com/xuri/excelize"
)

func main() {
	fmt.Println("Config...")
	//cfg := leggiCFG("conf.json")

	//xlsx, err := excelize.OpenFile(os.Args[1])
	xlsx, err := excelize.OpenFile("Prova.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	rows := xlsx.GetRows("Sheet1")

	for _, row := range rows {
		for _, cell := range row {

			fmt.Print(cell, "(", reflect.TypeOf(cell), ")", "\t")
		}
		fmt.Println()
	}

}
