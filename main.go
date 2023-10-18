package main

import (
	"encoding/csv"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"strconv"
	"strings"
	"sync"

	"github.com/xuri/excelize/v2"
)

const dataRange = 100

func main() {
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// Create a new sheet.
	_, err := f.NewSheet("go-vs-rust")
	if err != nil {
		panic(err)
	}
	// write headers
	headers := []string{"rows", "go", "rust"}
	f.SetSheetRow("go-vs-rust", "A1", &headers)
	// create folders and data
	for i := 1; i <= dataRange; i++ {
		rows := i * 100
		// folder path
		folderPath := fmt.Sprintf("./data/%d", rows)
		// get cell name
		cell, err := excelize.CoordinatesToCellName(1, i+1)
		if err != nil {
			panic(err)
		}
		// set rows of col 1
		f.SetCellInt("go-vs-rust", cell, i*100)
		// make folder
		_ = os.Mkdir(folderPath, 0755)
		// make csv file
		csvFile, err := os.Create(fmt.Sprintf("%s/simple.csv", folderPath))
		if err != nil {
			panic(err)
		}
		// get writer
		csvWriter := csv.NewWriter(csvFile)
		// write herder
		csvWriter.Write([]string{"id", "name", "age"})
		for j := 0; j < rows; j++ {
			csvWriter.Write([]string{fmt.Sprintf("%d", j), fmt.Sprintf("name %d", j), fmt.Sprintf("age %d", j)})
		}
		csvWriter.Flush()
	}
	// test go binary
	runChildProcess(f, "./bin/csv-to-excel-go.exe", 2)
	// test rust binary
	runChildProcess(f, "./bin/csv-to-excel-rust.exe", 3)
	f.SaveAs("simple.xlsx")
}

func runChildProcess(f *excelize.File, binPath string, col int) {
	var wg sync.WaitGroup
	for i := 1; i <= dataRange; i++ {
		wg.Add(1)
		go func(i int) {
			defer wg.Done()

			rows := i * 100
			// folder path
			simpleDataPath, err := filepath.Abs(fmt.Sprintf("./data/%d/", rows))
			if err != nil {
				panic(err)
			}
			// execute binary
			res, err := exec.Command(binPath, simpleDataPath).Output()
			if err != nil {
				panic(err)
			}
			passedTime := strings.TrimSpace(string(res))
			fmt.Println(passedTime)
			cell, err := excelize.CoordinatesToCellName(col, i+1)
			if err != nil {
				panic(err)
			}
			passedTimeInMs := 0.0
			// with time in ms
			if strings.HasSuffix(passedTime, "ms") {
				passedTime = strings.TrimSuffix(passedTime, "ms")
				parsedTime, err := strconv.ParseFloat(passedTime, 32)
				if err != nil {
					panic(err)
				}
				passedTimeInMs = parsedTime
			}

			// with time in s
			if strings.HasSuffix(passedTime, "s") {
				passedTime = strings.TrimSuffix(passedTime, "s")
				parsedTime, err := strconv.ParseFloat(passedTime, 32)
				if err != nil {
					panic(err)
				}
				passedTimeInMs = parsedTime * 1000
			}
			// set rows of col 1
			f.SetCellFloat("go-vs-rust", cell, passedTimeInMs, 5, 32)

		}(i)
	}
	// Wait for all child processes to complete
	wg.Wait()
}
