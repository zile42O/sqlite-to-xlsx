package main

import (
	"database/sql"
	"fmt"
	"log"
	"os"
	"strings"
	"time"

	"github.com/fatih/color"
	_ "github.com/mutecomm/go-sqlcipher/v4"
	"github.com/tealeg/xlsx"
)

func main() {
	startTime := time.Now()

	color.Red("\n\n----------------------")
	color.Cyan("Starting...")

	if len(os.Args) < 2 {
		color.Red("Please input the path of .db file")
		return
	}
	dbPath := strings.Join(os.Args[1:], " ")

	dbName := getFilenameWithoutExtension(strings.TrimPrefix(dbPath, "file:"))

	color.Cyan("Generating files for db: %s", dbPath)

	db, err := sql.Open("sqlite3", dbPath)
	if err != nil {
		log.Fatal(err)
	}
	defer db.Close()

	tables, err := getTableNames(db)
	if err == nil {
		for _, table := range tables {
			rows, err := db.Query("SELECT * FROM " + table)
			if err != nil {
				continue
			}
			defer rows.Close()

			file := xlsx.NewFile()
			sheet, err := file.AddSheet(table + "-sheet")
			if err != nil {
				continue
			}

			columns, err := rows.Columns()
			if err != nil {
				continue
			}

			headerRow := sheet.AddRow()
			for _, column := range columns {
				cell := headerRow.AddCell()
				cell.SetString(column)
			}

			for i := range columns {
				sheet.SetColWidth(i, i, 12)

				maxLen := len(columns[i])
				rows, err := db.Query("SELECT " + columns[i] + " FROM " + table)
				if err != nil {
					continue
				}
				defer rows.Close()
				for rows.Next() {
					var value string
					if err := rows.Scan(&value); err != nil {
						continue
					}
					if len(value) > maxLen {
						maxLen = len(value)
					}
				}

				if maxLen > 12 {
					sheet.SetColWidth(i, i, float64(maxLen))
				}
			}

			values := make([]interface{}, len(columns))
			for i := range values {
				values[i] = new(interface{})
			}
			for rows.Next() {
				if err := rows.Scan(values...); err == nil {
					row := sheet.AddRow()
					for _, value := range values {
						cell := row.AddCell()
						cell.SetValue(*value.(*interface{}))
					}
				}
			}
			err = file.Save("out/" + dbName + "_" + table + ".xlsx")
			if err != nil {
				color.Red("Failed to save file: %s", err)
			}
			color.Green("Excel file for table '%s' successfully created.\n", table)
		}
	} else {
		color.Red("Failed to get tables")
	}
	color.Cyan("Program finished..")
	color.Red("----------------------")
	color.Cyan("\n\nBy Zile42O")
	endTime := time.Now()
	elapsedTime := endTime.Sub(startTime)

	fmt.Printf("Program took: %d:%02d\n", int(elapsedTime.Minutes()), int(elapsedTime.Seconds())%60)
}

func getTableNames(db *sql.DB) ([]string, error) {
	rows, err := db.Query("SELECT name FROM sqlite_master WHERE type='table'")
	if err != nil {
		return nil, err
	}
	defer rows.Close()

	var tables []string
	for rows.Next() {
		var table string
		if err := rows.Scan(&table); err != nil {
			return nil, err
		}
		tables = append(tables, table)
	}
	return tables, nil
}

func getFilenameWithoutExtension(path string) string {
	parts := strings.Split(path, "/")
	filename := parts[len(parts)-1]
	if strings.Contains(filename, "?") {
		filename = filename[:strings.Index(filename, "?")]
	}
	return filename
}
