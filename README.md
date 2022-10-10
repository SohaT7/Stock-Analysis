# Stock Analysis
## Table of Contents
- [Overview of the Analysis](#overview-of-the-analysis)
    - [Purpose](#purpose)
    - [About the Dataset](#about-the-dataset)
    - [Tools Used](#tools-used)
    - [Description](#description)
- [Results](#results)
    - [Comparing the Stock Performance between 2017 and 2018](#Comparing-the-Stock-Performance-between-2017-and-2018)
    - [Execution times of the Original script and the Refactored script](#Execution-times-of-the-Original-script-and-the-Refactored-script)
- [Summary](#summary)
- [Contact Information](#contact-information)

## Overview of the Analysis
### Purpose:
The purpose of [this analysis](https://github.com/SohaT7/Stock-Analysis/blob/main/VBA_Challenge.xlsm) initially was to help the clients decide whether they would like to invest in the DQ stock (a green energy stock that interested them). The performance of the DQ stock based on the yearly return in 2018 was in the negative (it fell by 62.6%), and so investing in that was out of the question. The analysis was thus expanded to include gauging the performace of the other 11 stocks, in the years 2017 and 2018 both, from which the clients could then decide wherein to invest. 

### About the Dataset:
The dataset comprises 3012 records (rows) containing information on 12 different stocks (a record pertains to each day the stock opened and closed) in the following Excel file:

[Stock Data](https://github.com/SohaT7/Stock-Analysis/blob/main/green_stocks.xlsm)

### Tools Used:
 - VBA (Visual Basic for Applications)

### Description:
![Input Box](https://github.com/SohaT7/Stock-Analysis/blob/main/Images/InputBox_YearValue.png)

This analysis aims to gauge the performance of multiple stocks in the years 2017 and 2018. The code takes in, as input, the year value the client is looking to run the analysis for. It then calculates the Total Daily Volume (the frequency with which a stock is shared in a given day) and the Yearly Return (the percentage change in the value of the stock from the beginning to the end of the year) for all the 12 stocks in that year. 

For this analysis, a code was written which was then refactored to help run more efficiently. The buttons in our sheet help run both these codes (the refactored code runs faster).

![Initial](https://github.com/SohaT7/Stock-Analysis/blob/main/Images/Initial.png)

## Results
### Comparing the Stock Performance between 2017 and 2018:
![Output 2017](https://github.com/SohaT7/Stock-Analysis/blob/main/Images/Output_OriginalCode_2017.png)

![Output 2018](https://github.com/SohaT7/Stock-Analysis/blob/main/Images/Output_OriginalCode_2018.png)

Most of the stocks in 2017 had a positive performance i.e. a positive yearly return except for TERP which fell by 7.2%. The highest performing stock in 2017 was DQ with a yearly return of 199.4%, followed by SEDG with a return of 184.5%. The only other stocks that yielded a return greater than 100% in 2017, were ENPH (129.5%) and FSLR (101.3%). Even though the TERP stock had a total daily volume of $139,402,800, which is higher than that of the other four stocks (AY, DQ, HASI, and VSLR), it could not be used to gauge the value of the stock itself (but could gauge the yearly return). For instance, DQ which had a lower total daily volume of $35,796,200 had a yearly return that is much higher (199.4%).

By comparison, most of the stocks in 2018 performed in the negative, i.e. the yearly return for most stocks was in the negative. There is a positive return only on ENPH (81.9%) and RUN (84.0%). The lowest performing stocks in 2018 were DQ which fell by 62.6%, closely followed by JKS which fell by 60.5%. There were 3 other stocks that fell by more than 20%: FSLR, HASI, and SPWR. The total daily volume, again, was not reflective of the value of the stock, like yearly return was. Even though the total daily volume for ENPH is highest at $607,473,500 and so is its yearly return at 81.9%, this pattern does not follow for other stocks: the stock with the second highest total daily volume SPWR ($538,024,300) has the third lowest return that year (-44.6%). 

### Execution times of the Original script and the Refactored script:
The execution time for the original code is shown below:

![Time_Orinigal Code_2017](https://github.com/SohaT7/Stock-Analysis/blob/main/Images/ElapseTime_OriginalCode_2017.png)
![Time_Original Code_2018](https://github.com/SohaT7/Stock-Analysis/blob/main/Images/ElapseTime_OriginalCode_2018.png)

The execution time for the refactored code, by comparison, is shown below:

![Time_Refactored Code_2017](https://github.com/SohaT7/Stock-Analysis/blob/main/Images/VBA_Challenge_2017.png)
![Time_Refactored Code_2018](https://github.com/SohaT7/Stock-Analysis/blob/main/Images/VBA_Challenge_2018.png)

The refactored code runs faster than the original code. The execution time for the original code comes out to be around 0.6015625 seconds for both the year 2017 analysis run and the 2018 analysis run. However, the execution time for running the refactored code is shorter: it is around 0.1328125 seconds for both years' analysis run. However, the code (original as well as refactored) was run multiple times for the year 2017 and 2018 in order to get average or stable values, and the run time values stabilized after a couple times or so, except for the refactored code run on 2017 dataset which on a subsequent run gave an even shorter run time value of 0.125 seconds (it stabilized thereafter). This discrepancy can occur because on the first run especially, the computer has to allocate resources to run the macro which takes longer. However, once these resources have been allocated, they can readily be used and thus the code run then gives us a stable (and smaller) subsequent run time value. It is to be noted that the refactored code also includes the code for formatting within, and it still runs faster than the original code (which does not contain the code portion for formatting within it).

![Time_Refactored Code_2017_2](https://github.com/SohaT7/Stock-Analysis/blob/main/Images/ElapseTime_RefactoredCode_2017_2.png)

The difference in the original code and the refactored code is that in the former, we used the array of 12 tickers (for the 12 stocks) created to loop through all the rows in our dataset 12 times - each one run corresponds to one loop and by the end of the loop, the ticker was updated to the next one, so on and so forth. 

![Array for tickers](https://github.com/SohaT7/Stock-Analysis/blob/main/Images/Code_Initializing%20Array.png)

The Total Daily Volume, Starting Price, and Ending Price were extracted and displayed as output for each stock, in each run through the loop. 

![Original Code](https://github.com/SohaT7/Stock-Analysis/blob/main/Images/Code_Original%20Code.png)

In the refactored code, we created three additional arrays for Total Daily Volumes, Starting Prices, and Ending Prices. This helps run the code through all the rows in the dataset only once. In the one run through the loop, the variable tickerIndex updates itself to the next numeric value and thus updates the ticker value to the next one when it reaches the new ticker's rows in the dataset. In addition, when within a particular ticker's rows, the code collects and calculates values for Total Daily Volumes, Starting Prices and Ending Prices and stores them in their respective arrays. All this aids in running through the code once and conducting all the operations simultaneously. 

![Refactored Code](https://github.com/SohaT7/Stock-Analysis/blob/main/Images/Code_Refactored%20Code.png)

Later, the values in these arrays (tickers, total daily volumes, starting prices, and ending prices) are used to produce and display the output in the sheet. 

![Refactored_Output](https://github.com/SohaT7/Stock-Analysis/blob/main/Images/Code_Refactored_Output%20Arrays.png)

## Summary
Refactoring code makes the code more efficient, i.e. allows the code to have an even shorter run time, while keeping its functionality the same as before. It makes the code more efficient by reducing the number of coded lines, helping use up less space in the memory, or by refining the code's logic even further. Refactoring the code can make it more simple and easy to use, and even more flexible so that, if need be, other functionalities may be added with relative ease into it. The disadvantages of refactoring code include it being a potentially time-consuming process. It can also lead to the code becoming riddled with bugs in the process, which in turn may use up considerable time to debug and solve. 

While refactoring our original code, we ended up making it more efficient (the run time has decreased from around 0.62 seconds to 0.13 seconds) by breaking it down into more logical, neat, and concise steps: the code has to run through the dataset for a reduced number of times now and that saves time. A common element - the variable tickerIndex - is used to refer to all the arrays (including additional ones) we created which makes the code more neat and logical. While refactoring the code, it took some time to debug the bugs that appeared but once done, the code is highly efficient! 

## Contact Information
Email: st.sohatariq@gmail.com

