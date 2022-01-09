# Stock-Analysis

## Overview of the Project
### Background of the Project
The purpose of this project initially was to help Steve's parents decide whether they would like to invest in the DQ stock (a green energy stock they were interested in). The performance of the DQ stock based on the yearly return in 2018 was in the negative (it fell by 62.6%), and so investing in that was out of the question. The project was thus expanded to include gauging the performace of the other 11 stocks as well, from which Steve's parents can then decide wherein to invest. 
### Purpose and Description of the Project
![Input Box](https://github.com/SohaT7/Stock-Analysis/blob/main/InputBox_YearValue.png)

This project analyzes the performance of multiple stocks in the year 2017 and 2018. The code takes in the year value we are looking to run the analysis for and then calculates the Total Daily Volume (the frequency with which a stock is shared in a given day) and the Yearly Return (the percentage change in the value of the stock from the beginning to the end of the year) for all the 12 stocks in that year. 
For this project, a code was written which was then refactored to help run the program more efficiently. The buttons in our sheet help run both these codes (the refactored code runs faster).

![Initial](https://github.com/SohaT7/Stock-Analysis/blob/main/Initial.png)

## Results
### Comparing the Stock Performance between 2017 and 2018
![Output 2017](https://github.com/SohaT7/Stock-Analysis/blob/main/Output_OriginalCode_2017.png)

![Output 2018](https://github.com/SohaT7/Stock-Analysis/blob/main/Output_OriginalCode_2018.png)

Most of the stocks in 2017 had a positive performance i.e. a positive yearly return except for TERP which fell by 7.2%. The highest performing stock in 2017 was DQ with a yearly return of 199.4%, followed by SEDg with a return of 184.5%. The only other stocks that yielded a return greater than 100% in 2017 were ENPH (129.5%) and FSLR (101.3%). Even though the TERP stock has a total daily volume of $139,402,800 which is higher than that of 4 other stocks (AY, DQ, HASI, and VSLR), it can not be used to gauge the value of the stock as well as the yearly return can be. For instance, DQ which has a lower total daily volume of $35,796,200 has a yearly return that is much higher (199.4%).

By comparison, most of the stocks in 2018 performed in the negative, i.e. the yearly return for most stocks is in the negative. There is a positive return on only ENPH (81.9%) and RUN (84.0%). The lowest performing stocks in 2018 were DQ which fell by 62.6%, and closely followed by JKS which fell by 60.5%. There are 3 other stocks that fell by more than 20% (FSLR, HASI, SPWR). The total daily volume again is not reflective of the value of the stock like yearly return is. Even though the total daily volume for ENPH is highest at $607,473,500 and so is its yearly return at 81.9% the highest, this pattern does not follow: the stock with the second highest total daily volume SPWR ($538,024,300) has the third lowest return that year (-44.6%). 

### Execution times of the Original script and the Refactored script
The execution time for the original code is shown below:

![Time_Orinigal Code_2017](https://github.com/SohaT7/Stock-Analysis/blob/main/ElapseTime_OriginalCode_2017.png)
![Time_Original Code_2018](https://github.com/SohaT7/Stock-Analysis/blob/main/ElapseTime_OriginalCode_2018.png)

The execution time for the refactored code by comparison, is shown below:

![Time_Refactored Code_2017](https://github.com/SohaT7/Stock-Analysis/blob/main/VBA_Challenge_2017.png)
![Time_Refactored Code_2018](https://github.com/SohaT7/Stock-Analysis/blob/main/VBA_Challenge_2018.png)

The refactored code runs faster than the original code. The excution time for the original code come sout to be around 0.6015625 seconds for both the year 2017 analysis run and the 2018 analysis run. However, the execution time for running the refactored code is shorter: it is around 0.1328125 seconds for both years' analysis run. However, the code (original as well as refactored) was run multiple times for the year 2017 and 2018 in order to get average or stable values. The run time values stabilize after a couple times or so, except for the refacotred code run on 2017 dataset which on a subsequent run gave an even shorter run time value of 0.125 seconds. This discrepancy can occur because on the first run especially the computer has to allocate resources to run the macro which takes longer. However, once these resources have been allocated, they can readily be used and thus give us a stable (and smaller) subsequent run time value.

![Time_Refactored Code_2017_2](https://github.com/SohaT7/Stock-Analysis/blob/main/ElapseTime_RefactoredCode_2017_2.png)

The difference in the original code and the refactored code is that in the former, we created an array of tickers for the 12 stocks and then for each ticker/stock looped through all the rows in our dataset to get the Total Daily Volume, Starting Price, and Ending Price (the latter two were used to then calculate the Yearly Return before dislaying it). This means that the code ran through the entire code 12 times, one time each for each ticker/stock in our dataset.

The refactored code, on the other hand, runs through the entire code only once, thereby cutting down on the execution time considerably. It does so by creating three other arrays in addition to the array for tickers (arrays for total daily volumes, starting prices, and ending prices). As it runs through all the rows in the dataset, it keeps extracting and calculating the needed values and placing them in their respective arrays. After the one and only run, it displays the values from the arrays as is, except for those in the arrays for starting prices and ending prices, which it does a calcualtion upon before finally dispalying the Yearly Return value (shown below in the code). 

In the refactored code, creating a variable called tickerIndex and increasing its value by 1 
## Summary
