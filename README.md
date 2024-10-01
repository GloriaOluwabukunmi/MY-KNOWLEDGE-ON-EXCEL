### MY-KNOWLEDGE-ON-EXCEL
## MICROSOFT EXCEL
Microsoft Excel is a powerful spreadsheet program that is part of the Microsoft Office suite. It allows users to organize, format, and calculate data using a system of rows and columns
[Click_here](www.microsoft.com)

- ### SOME EXCEL FUNCTIONS
 1. SUM
 2. AVERAGE
 3. MIN
 4. MAX
 5. COUNT, COUNTA
 6. SEARCH, FIND
 7. RAND, RANDARRAY, RANDBETWEEN

### CONDITIONAL FUNCTIONS
- IF, IFS
- SWITCH
- AND
- OR, XOR
- SUMIF, SUMIFS
- MINIF, MINIFS
- MAXIF, MAXIFS
- AVERAGEIF
  
### Primary functions used for extracting texts in excel
- LEFT
- MID
- RIGHT

### TEXT CLEANING
Text cleaning is the process of preparing and refining text data for analysis which involves removing or correcting elements that could interfere with data processing or analysis which includes;
- Removing Noise
- Lowercasing: Converting all text to lowercase to ensure uniformity
- Correcting Typos: Fixing misspellings or grammatical errors to ensure consistency.
- Removing Duplicates: Ensuring that repeated entries are consolidated.
- Handling Missing Values: Deciding how to deal with empty or null values, either by filling them in or removing the entries altogether.
  
- ### Some of the functions used for Data Cleaning
  1. LOWER
  2. UPPER
  3. PROPER

### NESTING
the practice of placing one function inside another function, it allows more complex calculations and data manipulation
example of nested function
=IF (J8 <=20, "LOW", IF(J8 <=50, "MEDIUM", "HIGH"))

### LOOKUP FUNCTIONS
the LOOKUP function is used to search for a value in a range and return a corresponding value from another range

-Some LOOKUP Functions

1. - ### VLOOKUP;
Specifically designed for vertical lookups

VLOOKUP (lookup_value, table_array, col_index_num, [range_lookup])

Table_array: The range that contains the data.

Col_index_num: The column number in the table from which to retrieve the value.

Range_lookup: TRUE for an approximate match or FALSE for an exact match.

2. - ### HLOOKUP
Similar to VLOOKUP, but it is specifically designed for horizontal rows:
HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])

3. INDEX
4. MATCH
5. XLOOKUP
6. CHOOSE
7. INDIRECT

### Data Validation
It controls the data that enters the cells
Ensure data integrity by setting rules for data entry

#### Shortcutt for Data Validation ALT+AVV

- Types of Data Validation
- Whole Number
- Decimal
- List
- Date
- Time
- Text length
- Custom

  ### PIVOT TABLE
  Pivot table is a powerful excel tool used for data summarization

  It allows quick analysis of large dataset inorder to extract meaningful insights without needing to write complex formulas
   ### some shortcut keys for pivot table:
  - ctrl+shift+1 used to sort numbers
     e.g 123456789 + 123,456,789
    - Alt+F1 or Alt+F1+Fn to get a chart on the pivot table
    - ctrl+X to cut the chart on the pivot table
    - ctrl+V to paste the chart

  
  

