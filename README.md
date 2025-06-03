# Match and Color Log Groups in Excel

This R script processes a logbook Excel file by grouping related questions, assigning colors to each group, and writing the results back to Excel with formatting.

## Features

- Reads two Excel files: 
  - A log file (with a sheet named `02_Logbook`)
  - A checks file (with `question.name` column, possibly containing comma-separated values)
- Groups log entries based on `uuid` and `question.name` matches
- Applies color formatting to each group in the output Excel
- Saves the result to a new Excel file

## Requirements

- R
- The following R packages:
  - `dplyr`
  - `tidyr`
  - `stringr`
  - `openxlsx`

## How to Use

1. Place your Excel files in an `input` folder:
   - `Master_logs.xlsx` (must include a sheet named `02_Logbook`)
   - `checks.xlsx` (must include a `question.name` column)

2. Run the R function:

```r
process_logs("./input/Master_logs.xlsx", "./input/checks.xlsx")


The result will be saved in ./output/Master_logs.xlsx with grouped and color-highlighted log entries.

License
MIT License
