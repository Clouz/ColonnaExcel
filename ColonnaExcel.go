package main

import (
	"fmt"
	"strconv"

	"github.com/xuri/excelize"
)

func main() {
	fmt.Println("Config...")
	cfg := leggiCFG("conf.json")

	//xlsx, err := excelize.OpenFile(os.Args[1])
	xlsx, err := excelize.OpenFile("Prova.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	rows := xlsx.GetRows("Sheet" + strconv.Itoa(cfg.Pagina))

	//Head
	fmt.Print(cfg.NomeRange, "\t")
	for index := 0; index < len(cfg.ColonneRipetute); index++ {
		fmt.Print(cfg.ColonneRipetute[index].Intestazione, "\t")
	}
	fmt.Println()

	//Row
	for index := cfg.RigaIniziale - 1; index < len(rows); index++ {

		//Collumn
		for iRow := cfg.RangeStart - 1; iRow < cfg.RangeStop; iRow++ {
			fmt.Print(rows[index][iRow], "\t")

			//extra collumn
			for _, extra := range cfg.ColonneRipetute {
				fmt.Print(rows[index][extra.Colonna-1], "#\t")
			}
			fmt.Println()
		}
	}
}
