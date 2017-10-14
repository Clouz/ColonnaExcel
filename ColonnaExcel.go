package main

import (
	"bufio"
	"fmt"
	"os"
	"strconv"

	"github.com/cheggaaa/pb"
	"github.com/xuri/excelize"
)

var result [][]string

func main() {

	fmt.Println("Leggo Config...")
	cfg, err := leggiCFG("conf.json")
	if err != nil {
		fmt.Println("Config error: ", err)
		return
	}

	if len(os.Args) > 1 {
		fmt.Println("Leggo Excel...")

		result, err = LeggiExcel(os.Args[1], cfg)
		if err != nil {
			fmt.Println("Open File error: ", err)
			return
		}
	} else {
		fmt.Println("Impostare il file conf.json e successivamente trascinare una file excel su questo eseguibile")
		return
	}

	fmt.Println("Scrivo Excel...")
	err = ScriviExcel("result.xlsx", result, cfg)
	if err != nil {
		fmt.Println("Write FIle error: ", err)
		return
	}

	fmt.Println("Operazione completata, premere un tasto per chiudere...")
	bufio.NewReader(os.Stdin).ReadBytes('\n')
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

	//Render progress bar
	bar := pb.StartNew(ro)

	//Head
	result[0] = make([]string, co)
	result[0][0] = cfg.NomeRange

	for index := 0; index < len(cfg.ColonneRipetute); index++ {
		result[0][index+1] = cfg.ColonneRipetute[index].Intestazione
	}
	bar.Increment()

	i := 1
	//Row
	for index := cfg.RigaIniziale - 1; index < len(rows); index++ {
		//Collumn
		for iRow := cfg.RangeStart - 1; iRow < cfg.RangeStop; iRow++ {

			bar.Increment()
			exit := false
			row := rows[index][iRow]
			for _, skip := range cfg.CelleDaEscludere {
				if row == skip {
					exit = true
					break
				}
			}

			if !exit {
				ii := 0

				result[i] = make([]string, co)
				result[i][ii] = row

				//extra collumn
				for _, extra := range cfg.ColonneRipetute {
					ii++
					result[i][ii] = rows[index][extra.Colonna-1]
				}
				i++
			}
		}
	}

	bar.FinishPrint("Lettura Excel Terminata")

	return result, nil
}

// ScriviExcel dato un array multidimensionale restituisce un foglio excel scritto
func ScriviExcel(path string, data [][]string, cfg *Configuration) error {

	foglio := "Sheet1"

	xlsx := excelize.NewFile()
	//index := xlsx.NewSheet(foglio)

	//Render progress bar
	bar := pb.StartNew(len(data))

	for i, riga := range data {
		bar.Increment()
		for ii, cella := range riga {
			if ii == 0 && cella == "" {
				break
			}
			xlsx.SetCellValue(foglio, indexToAxis(i, ii), cella)
		}
	}

	bar.FinishPrint("Scrittura Excel Terminata")

	//xlsx.SetActiveSheet(index)
	err := xlsx.SaveAs(path)
	if err != nil {
		return err
	}

	return nil
}

func indexToAxis(row int, col int) string {
	var arr = []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ"}
	return arr[col] + strconv.Itoa(row+1)
}
