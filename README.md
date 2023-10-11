# 02_VBA_Challenge

## Overview

This repository contains a VBA script for analyzing and summarizing stock data from multiple years. The directions for the module had somewhat contradictory instructions: the solution image showed no conditional formatting for Percentage Change, but the instructions listed conditional formatting for Percentage Change. As a result, after reviewing the directions with the BCS team I decided to submit my solution as shown in the module and write an explanation of the solution that would work should there have been a need to also have conditional formatting in the Percentage Change column.  

Additionally, the excel file has several factors I wish to address here in the ReadMe. First, I did not use the listed preferred color green (RGB 0, 255, 0). This was intentional. I instead used the same shade of green (and red) from my previous homework. I realize this is a full portfolio of work and I would like my portfolio to show a congruent use of colors and styles across each of my pieces where possible. 

Lastly, I have exported the script module from it's origial name and renamed it outside of the excel file for clarification. It may require the file to be reloaded with the new name and run again. This should not alter the results of the script. 

## Alternate Code Explanation

If the directions and photos were wrong and the Percent Change column needed to be conditionally formatted based on colors, my solution would be (Applied in the appropriate location):

   If ws.Cells(Output_Row, PERCENT_CHG_COL).Value > 0 Then
       ws.Cells(Output_Row, PERCENT_CHG_COL).Interior.Color = RGB(51, 153, 51)
   ElseIf ws.Cells(Output_Row, PERCENT_CHG_COL).Value < 0 Then
       ws.Cells(Output_Row, PERCENT_CHG_COL).Interior.Color = RGB(255, 0, 0)
   Else
       ws.Cells(Output_Row, PERCENT_CHG_COL).Interior.Color = RGB(255, 255, 255)
   End If

The above result has been tested and has been found to work following calculations for format of the percentage of the percent change column. An alternate Dim could have been used to make the solution cleaner as well. 

## Results

Three images are included in the repository, each representing the results for a different year: 2018, 2019, and 2020.

## Usage

You can use these scripts to analyze stock data in Excel. Depending on the requirements, choose the appropriate script based on whether Conditional Formatting for Percentage Change is needed or not.

1. Open the respective Excel spreadsheet (`Tia Scott_VBA Challenge 02.xlsm).

2. Enable macros if prompted.

3. Run the VBA script (`TiaScott_VBA Challenge 02.bas) within Excel to perform the analysis and generate summary tables.

4. The results will be displayed in the Excel spreadsheet across all entire workbook.
 
