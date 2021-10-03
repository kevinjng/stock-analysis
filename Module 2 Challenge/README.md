# Stock Analysis

## Overview of Project
The purpose of this project is to utilize Excel, Visual Basic code (VBA) and analyze a dataset of twelve different stock tickers for the years 2017 & 2018. Datasets for both years include the ticker abbreviation, date, opening price, high & low prices, close & adjusted close prices, and the trading volume for the respective period.

By using VBA and applying it to worksheet 'All Stocks Analysis', a straightforward analysis is created by taking the ticker abbreviation and total daily volume, and calculating the return obtained. The VBA code for this worksheet loops through all the rows in the raw 2017 & 2018 datasets and references variables such as *tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices* in order to execute the calculation.

## Results

  
![2017_Stock Analysis Results](https://user-images.githubusercontent.com/90368828/135771724-2145325b-e917-4722-94fa-f49ac5ee8a9d.png)
![2018_Stock Analysis Results](https://user-images.githubusercontent.com/90368828/135771728-30278fc4-a607-419c-b050-ca3e5d1c692c.png)


After applying the analysis to both 2017 & 2018, 2017 shows more positive and marginal return compared to the output 2018 provides. There are significant returns with tickers such as DQ, ENPH, FSLR, and SEDG; providing 199.4%, 129.5%, 101.3% and 184.5% in returns respectively. The remaining tickers all provide some level of positive returns, with the only exception being TERP which had a negative return of (7.2)%.

As for 2018, almost all tickers provided negative returns besides ENPH and RUN with positive returns of 81.9% and 84.0%, respectively. Although DQ saw a significantly large, positive return in 2017, a different picture can be seen in 2018 where it produces the largest negative return of (62.6%). Tickers JKS, SPWR, and FSLR follow the trend of negative returns with (60.5%), (44.6%), and (39.7%), respectively.

## Summary
Between the original workbook, *green_stocks*, and the *VBA_Challenge* workbook with refactored code, I was able to shave off 0.300781 and 0.304688 seconds with the refactored VBA script for the 2017 & 2018 analysis, respectively. Comparison of script runtimes shown below.


**Original Runtime (left) and Refactored Runtime (right) for 2017 analysis**
![VBA_Challenge_2017_original](https://user-images.githubusercontent.com/90368828/135771734-b9ebc042-ef81-4930-b476-30d6c89c715c.png)
![VBA_Challenge_2017](https://user-images.githubusercontent.com/90368828/135771736-e4745045-933f-4b38-bc59-dbc8a8ff6f5d.png)


**Original Runtime (left) and Refactored Runtime (right) for 2018 analysis**
![VBA_Challenge_2018_original](https://user-images.githubusercontent.com/90368828/135771743-8022851b-bbd7-47b7-b624-2050d6087057.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/90368828/135771746-962af4c2-f344-4b32-b5a2-0a448ea6faff.png)

Upon review of the original and refactored scripts, there were 2 main contributors to the decreased runtimes. One of the first factors that assisted in shaving off milliseconds would be the inclusion of the conditional formatting lines.

**“'Code for Conditional Formatting.
Worksheets("All Stocks Analysis").Activate
Range("A3:C3").Font.Bold = True
Range("C4:C15").Font.Bold = True
Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("B4:B15").NumberFormat = "#,##0"
Range("C4:C15").NumberFormat = "0.0%"
Columns("B").AutoFit

dataRowStart = 4
dataRowEnd = 15
For I = dataRowStart To dataRowEnd

 If Cells(I, 3) > 0 Then
   'Color the cell green
   Cells(I, 3).Interior.Color = vbGreen

 ElseIf Cells(I, 3) < 0 Then

   'Color the cell red
   Cells(I, 3).Interior.Color = vbRed
    
 Else
    'Clear the cell color
    Cells(I, 3).Interior.Color = xlNone
    
 End If
Next I”**

In the original workbook, I had this portion of script apportioned to a completely different macro that would run after the *AllStocksAnalysis* script. But when refactored, I included it within the same macro near the ending lines of script. I believe that the total amount of macros in the workbook were reduced by including this portion in the main macro, therefore reducing the total runtime.

Moving on to the second and final factor, I believe the positioning of the end-timer code played a role in streamlining the entire macro (portion of code shown below).

**“endTime = Timer
MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)”**

For both original and refactored workbooks, this portion of script was placed at the very end of the macro right before the ‘End Sub’ line. When I originally began refactoring, I just did the inclusion of the conditional formatting lines but initially left the end-timer lines where they originally were. When I ran the analysis with the script oriented this way, I noticed there was only a slight significant decrease in total runtime, and that the analysis was not running smooth as possible.

After noticing this, I realized where the hiccup was in the macro, and placed the end-timer lines right before the ‘End Sub’ command. When running the macro with the refactored script, I noticed a much more noticeable decrease in execution time and, the macro ran much smoother with the end-timer lines in the correct position.

Overall, refactoring code presents advantages as well as some disadvantages in regard to certain aspects. To begin with, refactoring is advantageous in the fact that it is optimizing your macro, program, etc. and streamlining the process which creates a faster and efficient running program. This was shown with refactoring the stock analysis code but can be portrayed more effectively in real life.

Take video games for example, when the source code is originally written for a game, there are most likely portions of code that are inadvertently long or possibly complicate the layout and understanding of the code. Not only would it make it harder to understand for fellow developers working on the game, but it would most definitely create a very slow-running version of the game that would not be enjoyable by consumers. Refactoring the source code would allow for the game engine to run much smoother, and would be less intensive on the CPU, GPU, and RAM for both the developers and end-consumer.

Simultaneously, refactoring code does introduce some disadvantages. For one, refactoring certain portions could possibly require significantly deeper understanding of the specific language which would lengthen the overall development time of the video game. This deeper understanding could potentially lead to more complex strings of code being written which will optimize the overall program but could lead to potential complications when novice developers are editing or adding further code.

In conclusion, refactoring code produces both advantages and disadvantages, but in the long run and for consumer use, the advantages should outweigh the disadvantages since a more streamlined and optimal program is created as a result.

