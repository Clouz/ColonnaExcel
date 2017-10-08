package main

import (
	"encoding/json"
	"fmt"
	"os"
)

// Configuration represent conf.json
type Configuration struct {
	Pagina          int
	RigaIniziale    int
	RangeStart      int
	RangeStop       int
	NomeRange       string
	ColonneRipetute []struct {
		Intestazione string
		Colonna      int
	}
}

func leggiCFG(cfgName string) (*Configuration, error) {

	file, err := os.Open(cfgName)
	if err != nil {
		return nil, err
	}
	decoder := json.NewDecoder(file)
	Conf := Configuration{}
	err = decoder.Decode(&Conf)
	if err != nil {
		return nil, err
	}

	fmt.Print("Sheet: ", Conf.Pagina, " ## Start: ", Conf.RangeStart, " ## Stop: ", Conf.RangeStop, " ## [")
	for _, num := range Conf.ColonneRipetute {
		fmt.Print(num, "; ")
	}
	fmt.Println("]")

	return &Conf, err
}
