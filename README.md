# Old EPPlus (4.5.3.2) vs New EPPlus (7.x)

This repo contains two independent test projects to highlight the issues I've found with the new EPPlus, which is blocking me from upgrading 4.5.3.2 to v7.

# 1. OpenFileWithConditionalFormattingTest

This test attempts to open a the "YourResultsPart.xlsm" file with some conditional formatting in it, which opens fine in 4.5.3.2 but now throws an exception. 

The error is in ExcelConditionalFormattingDataBar but I couldn't figure out how to fix it (changing Double.Parse to Double.TryParse for HighValue and LowValue just leads to errors elsewhere).

# 2. CalculationTest

This test demonstrates an issue with the Calculate() method in new EPPlus (it works fine in old EPPlus). 
After changing two input cells, only one of them gets recalculated. 
I think there must be an issue with the recalculation tree, in terms of keeping track of which formula cells are "dirty" and need to be recalculed.

Note that I couldn't replicate this issue when starting from a new empty spreadsheet (see NewExcelFileCalculationTest), but in reality I never work with empty spreadsheets...