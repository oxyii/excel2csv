# excel2csv - Convert Excel files to CSV files

with headers search

Useful for parsing incoming price lists, etc.
Can read both XLS and XLSX files.

## Use case

Let's say you know exactly which columns should be in each file:

```go
var (
    ColumnBrand = "brand"
    ColumnDesc  = "name"
    ColumnPrice = "price"
)
```

but you're not sure how exactly they're titled. Let's make a map of possible column names:

```go
func GetPossibleHeaders(filename string) map[string]string {
    switch filename {
    case "Supplier1_file1.xlsx":
        return map[string]string{
            "make":         ColumnBrand,
            "manufacturer": ColumnBrand,
            "description":  ColumnDesc,
            "price USD":    ColumnPrice,
        }
    case "Supplier2_file13.xls":
        return map[string]string{
            "manufacturer":   ColumnBrand,
            "name_en":        ColumnDesc,
            "description_en": ColumnDesc,
            "price day":      ColumnPrice,
        }
    
        // other cases
        default:
            panic("unknown file")
    }
}
```

In function above we're mapping possible column names to the ones we know.
Possible column name can be a substring of the actual column name.

Now we can scan the directory for files and convert them to CSV:

```go
package main

import (
    "os"

    "github.com/oxyii/excel2csv"
)

func ConvertSheet(filename string, sheet *excel2csv.Sheet) error {
    headers := GetPossibleHeaders(filename)

    outputFilename := "output/" + filename + "_" + sheet.Name + ".csv"
    outputFile, err := os.Create(outputFilename)
    if err != nil {
        return err
    }
    
    defer outputFile.Close()
    
    if err := sheet.Convert(outputFile, headers, []string{ColumnPrice}); err != nil {
        return err
    }
    
    return nil
}

func main() {
    files, err := os.ReadDir("files")
    if err != nil {
        panic(err)
    }
    
    for _, file := range files {
        if file.IsDir() {
            continue
        }
    
        filename := file.Name()
        sheets, err := excel2csv.Open("files/" + filename)
        if err != nil {
            panic(err)
        }
    
        for _, sheet := range sheets {
            if err := ConvertSheet(filename, sheet); err != nil {
                panic(err)
            }
        }
    }
}
```