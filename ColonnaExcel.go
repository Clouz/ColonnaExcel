package main

import (
	"fmt"
	"strconv"

	"github.com/xuri/excelize"
)

func main() {
	fmt.Println("Config...")
	cfg := leggiCFG("conf.json")
	result, err := LeggiExcel("Prova.xlsx", cfg)
	if err != nil {
		fmt.Println(err)
		return
	}
	err = ScriviExcel("result.xlsx", result, cfg)
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Print(result)
	//xlsx, err := excelize.OpenFile(os.Args[1])
}

// LeggiExcel Legge il contenuto di un excel e lo incolonna
func LeggiExcel(path string, cfg *Configuration) ([][]string, error) {

	xlsx, err := excelize.OpenFile(path)
	if err != nil {
		return nil, err
	}

	rows := xlsx.GetRows("Sheet" + strconv.Itoa(cfg.Pagina))

	ro := 1 + (len(rows)-cfg.RigaIniziale+1)*(cfg.RangeStop+1-cfg.RangeStart)
	co := 1 + len(cfg.ColonneRipetute)
	result := make([][]string, ro)

	//Head
	result[0] = make([]string, co)
	result[0][0] = cfg.NomeRange
	fmt.Print(result[0][0], "\t")

	for index := 0; index < len(cfg.ColonneRipetute); index++ {
		result[0][index+1] = cfg.ColonneRipetute[index].Intestazione
		fmt.Print(result[0][index+1], "\t")
	}
	fmt.Println()

	i := 1
	//Row
	for index := cfg.RigaIniziale - 1; index < len(rows); index++ {

		//Collumn
		for iRow := cfg.RangeStart - 1; iRow < cfg.RangeStop; iRow++ {
			ii := 0
			result[i] = make([]string, co)

			result[i][ii] = rows[index][iRow]
			fmt.Print(result[i][ii], "\t")

			//extra collumn
			for _, extra := range cfg.ColonneRipetute {
				ii++
				result[i][ii] = rows[index][extra.Colonna-1]
				fmt.Print(result[i][ii], "#\t")
			}
			i++
			fmt.Println()
		}
	}

	return result, nil
}

// ScriviExcel dato un array multidimensionale restituisce un foglio excel scritto
func ScriviExcel(path string, data [][]string, cfg *Configuration) error {

	foglio := "Sheet1"

	xlsx := excelize.NewFile()
	index := xlsx.NewSheet(foglio)

	for i, riga := range data {
		for ii, cella := range riga {
			xlsx.SetCellValue(foglio, indexToAxis(i, ii), cella)
		}
	}

	xlsx.SetActiveSheet(index)
	err := xlsx.SaveAs(path)
	if err != nil {
		return err
	}

	return nil
}

func indexToAxis(row int, col int) string {
	var arr = []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

	return arr[col] + strconv.Itoa(row+1)
}
