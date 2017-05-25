# Annualization Custom Function for Excel

## What is it?
It's a function written in VBA for Excel to annualize income or expense data. When income or expense figures are reported in year-to-date terms, they must be "annualized" to convert them to end-of-year terms. This is a crucial step when calculating financial ratios such as return on assets, return on equity, or any ratio that compares income statement data to balance sheet data. 

## How does it work?
The function accepts 2 arguments: income (or expense) and date. The function will then figure out if the date supplied is in a leap year, then it will multiply the income (or expense) figure by a calculated "annualization coefficient" based on the supplied date.

## Usage
Import the `annualize.bas` file into a new module in the Developer window in Excel. Make sure macros are enabled. 

Set up a table with these fields and data:

|#  | A        | B             | C                  |
|---|:--------:|:-------------:|:------------------:|
|1  | Amount   | Date          | Annualized Amount  |
|2  | $100.00  | 5/25/2017     | `=Annualize(A2,B2)`| 


Cell C2 should now return $251.72.

|#  | A        | B             | C                  |
|---|:--------:|:-------------:|:------------------:|
|1  | Amount   | Date          | Annualized Amount  |
|2  | $100.00  | 5/25/2017     | $251.72            |
