### MY KNOWLEDGE ON MICROSOFT EXCEL

[MICROSOFT EXCEL](#microsoft-excel)

[EXCEL FUNCTIONS](#excel-functions)

[CONDITIONAL FUNCTIONS](#conditional-functions)

[PRIMARY FUNCTIONS USED FOR EXTRACTING TEXT IN EXCEL](#primary-functions-used-for-extracting-text-in-excel)

[DATA CLEANING](#data-cleaning)

[SOME FUNCTIONS USED FOR DATA CLEANING](#some-functions-used-for-data-cleaning)

[NESTING](#nesting)

[LOOKUP FUNCTIONS](#lookup-functions)

[DATA VALIDATION](#data-validation)

[Shortcut for Data Validation](#shortcut-for-data-validation)

[Types of Data Validation](#types-of-data-validation)

[PIVOT TABLE](#pivot-table)


[CHARTS CREATED USING PIVOT TABLE](#charts-created-using-pivot-table)



#### MICROSOFT EXCEL📖
Microsoft Excel is a powerful spreadsheet program that is part of the Microsoft Office suite. It allows organization, formating, and calculation of data using a system of rows and columns
[Click_here](www.microsoft.com)

- #### EXCEL FUNCTIONS
 ---------------------------
 1. SUM
 2. AVERAGE
 3. MIN
 4. MAX
 5. COUNT, COUNTA
 6. SEARCH, FIND
 7. RAND, RANDARRAY, RANDBETWEEN

#### CONDITIONAL FUNCTIONS
-------------------------------------------------------
- IF, IFS
- SWITCH
- AND
- OR, XOR
- SUMIF, SUMIFS
- MINIF, MINIFS
- MAXIF, MAXIFS
- AVERAGEIF
  
#### PRIMARY FUNCTIONS USED FOR EXTRACTING TEXT IN EXCEL
---------------------------------------------------------------
- LEFT
- MID
- RIGHT

#### DATA CLEANING 
---
Text cleaning is the process of preparing and refining text data for analysis which involves removing or correcting elements that could interfere with data processing or analysis. It includes;
- Removing Noise
- Lowercasing: Converting all text to lowercase to ensure uniformity
- Correcting Typos: Fixing misspellings or grammatical errors to ensure consistency.
- Removing Duplicates: Ensuring that repeated entries are consolidated.
- Handling Missing Values: Deciding how to deal with empty or null values, either by filling them in or removing the entries altogether.
  
- #### SOME FUNCTIONS USED FOR DATA CLEANING
  1. LOWER
  2. UPPER
  3. PROPER

#### NESTING
---
the practice of placing one function inside another function, it allows more complex calculations and data manipulation of data.
example of nested function
=IF (J8 <=20, "LOW", IF(J8 <=50, "MEDIUM", "HIGH"))

#### LOOKUP FUNCTIONS
the LOOKUP function is used to search for a value in a range and return a corresponding value from another range

-Some LOOKUP Functions

1.  #### VLOOKUP;
Specifically designed for vertical lookups

VLOOKUP (lookup_value, table_array, col_index_num, [range_lookup])

Table_array: The range that contains the data.

Col_index_num: The column number in the table from which to retrieve the value.

Range_lookup: TRUE for an approximate match or FALSE for an exact match.

2.  #### HLOOKUP
Similar to VLOOKUP, but it is specifically designed for horizontal rows:
HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])

3. INDEX
4. MATCH
5. XLOOKUP
6. CHOOSE
7. INDIRECT

### DATA VALIDATION
---
It controls the data that enters the cells
Ensure data integrity by setting rules for data entry

#### Shortcut for Data Validation
- ALT+AVV

#### Types of Data Validation
- Whole Number
- Decimal
- List
- Date
- Time
- Text length
- Custom

  #### DATA VISUALIZATION
  ---
  ##### PIVOT TABLE 
  Pivot table is a powerful excel tool used for data summarization

  It allows quick analysis of large dataset inorder to extract meaningful insights without needing to write complex formulas
   - some shortcut keys used for pivot table:
  1. ctrl+shift+1 used to sort numbers
     e.g 123456789 + 123,456,789
    2. Alt+F1 or Alt+F1+Fn to get a chart on the pivot table
    3. ctrl+X to cut the chart on the pivot table
    4. ctrl+V to paste the chart

- #### CHARTS CREATED USING PIVOT TABLE

 ![image](https://github.com/user-attachments/assets/b86a1bf7-f3a9-4680-8abd-5109ee3cf70a)

![image](https://github.com/user-attachments/assets/d8465382-1768-4518-b4e0-2b6d4bf57c30)

![image](https://github.com/user-attachments/assets/37f12f85-3c05-42d4-a89e-bd9de90bb802)

![image](https://github.com/user-attachments/assets/a33c115e-b593-4abf-9b7a-da1499e66e59)

![image](https://github.com/user-attachments/assets/8497df70-d223-4cf4-917b-1a1cf5474e02)
