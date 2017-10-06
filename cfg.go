package main

import (
	"encoding/json"
	"fmt"
	"os"
)

// Configuration represent conf.json
type Configuration struct {
	RangeStart      int
	RangeStop       int
	ColonneRipetute []int
}

func leggiCFG(cfgName string) *Configuration {

	file, _ := os.Open(cfgName)
	decoder := json.NewDecoder(file)
	Conf := Configuration{}
	err := decoder.Decode(&Conf)
	if err != nil {
		fmt.Println(err)
	}

	fmt.Print("Start: ", Conf.RangeStart, " ## Stop: ", Conf.RangeStop, " ## [")
	for _, num := range Conf.ColonneRipetute {
		fmt.Print(num, "; ")
	}
	fmt.Println("]")

	return &Conf
}
