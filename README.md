# Overview of the Project- Stock Analysis using VBA

We have analysed a data set of 12 dfferent, Green Energy Sector stocks for the years 2017 and 2018. Using VBA we have calculated the following indicators for a given stock:
1. **Total Volume** traded of a stock in the year.
2. **Starting Price** of the stock at the beginning of the year.
3. **Ending Price** of the stock at the end of the year.


```
'2b) Loop over all the rows in the spreadsheet for every ticker
        For i = 2 To RowCount
            
'script to check the column A for ticker value and calculate the tickerVolume
            If Cells(i, 1).Value = ticker Then

'3a) Increase volume for current ticker
               tickerVolumes(tickerIndex) = (tickerVolumes(tickerIndex) + Cells(i, 8).Value)
            End If

'3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i - 1, 1).Value <> ticker And Cells(i, 1) = ticker Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
        
'3c) check if the current row is the last row with the selected ticker and
'If the next row’s ticker doesn’t match, increase the tickerIndex.
                
            If Cells(i + 1).Value <> ticker And Cells(i, 1) = ticker Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            
'3d Increase the tickerIndex.
        Next i
    
```
By calculating these 3 variables, we have tried to analyse how much a given stock has been traded during the year and how much would have been our return on that stock if we would have bought it at the start of year and sold at the end of the same year.

## Performance Measurement ##
To measure the performance of this code we captured the **Start time**  and **End time** of execution and then counting the seconds taken to completely run it

```
  Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    .
    .
    .
    .
     endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```

This calculation was automated usong VBA and done through a macro, **AllStockAnalysis**. 

Now as part of this challenge, we have tried to improve the performance of the same code by using Arrays and more efficient use of the For loops.

## Use of Refactoring technique ##

To Refactore the code we have changed the way For loops were running and instead of storing the value of 3 calculations namely **Total Volume**, **Starting Price** and **Ending Price**  in 3 variables, we have used 3 arrays to store the value of 3 fields for each Stock. 

# Results:

## Analysis of Stock Performance YoY ##

1. Most of the Green Energy stocks **(11 out of 12)** gave a postive return in the year 2017. The **RoI(Return on Investment)** went as high as **199.4 % (DQ stock)**. 

---
<img src = ".\Resources\VBA_Challenge_2017.png"></img>
---

2. On the other hand, year 2018 was not as good for the Green Energy Stocks as we can see from the figure 2.

---
<img src = ".\Resources\VBA_Challenge_2018.png"></img>
---

**10** out of **12** stocks analysed, gave a negative return in year 2018. 

If we look at the **Total Daily Volume** column in figure 1 and figure 2, we can see that stock DQ has been traded almost **3** times **YoY**.

## Result of Refactoring the Code

After refactoring the code, the execution time for the macro is improved by 20% in terms of execution time.

### Execution time for Year 2018 before and after refactoring the code:

As we can see from the below table that after refactoring the execution time reduced by almost **10%**. Thus resulting in a faster execution of the macro. 

| Year   | Before Refactor| After Refactor|
| -------|:--------------:| -------------:|
| 2017    | 0.8125    | 0.7265 |
| 2018    | 0.8320    | 0.7265 |


---

# Summary #

## Advantages of Refactoring the Code
1. Refactoring allows us to keep the code bug free. As while doing the refactoring we may be able to find some underlying bugs.
2. Refactoring allows us to understand the code and utilize the resources in a more effficient manner by way of changing the Variable types and redesigning the loops structure.
3. Refactoring 

## Disadvantages of Refactoring the Code

1. While refactoring a legacy system, we may introduce some unintended bugs to the system. As usually the code is too large to consider it all at once.
2. Refactoring may not justify the resources(time and money) spent on it. The Cost Benefit ratio may not be favourable.

## Impact of Refactoring our Code ##

As we can see from above table that after refactoring, our execution time has increased by ~10%. So it has made a significant improvement in code execution time.

In refactored code we have used Arrays to store the value of 3 calculations for variables **Total Volume**, **Starting Price** and **Ending Price** and used these to print the outcome for all the 12 stocks. Also the output is done outside the For Loops in the refactored version.
Using Array has significantly helped in reducing the execution time as we know that An array is considered to be a homogenous collection of data. As a result for any purpose if a user wishes to store multiple values of a similar type then arrays can be used and utilized efficiently.


[Credit](https://www.educba.com/advantages-of-array/)