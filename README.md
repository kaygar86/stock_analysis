# stock_analysis

# Overview of Project

The purpose of this analysis was to provide our friend Steve with a tool to analyze stock returns. We created a VBA script that analyzed the returns of 12 different stocks for the years 2017 and 2018. The original code used a nested loop that looped through all 3013 rows for each of the 12 tickers. 

Steve really liked the tool we provided and wanted to potentially use it to analyze thousands of stocks, which would be a much larger data set so we required a script than ran more efficiently. We refactored the code by removing the nested loop and including a ticker index that enabled the code to only loop through the 3013 rows one time and still deliver the analysis of each ticker. Using VBA's timer feature in the script we captured and compared the run times of the original and refactored code. Our results are documented below. 


# Results: 

There was no change in the stock analysis results when we ran the original code and the refactored code. The purpose of refactoring the code was to improve efficiency and run time. 

In 2017 the majority of the stocks we analyzed showed a positive return, with a few showing a return exceeding 100%. The analysis of 2018 showed the majority of stocks with negative returns. ENPH and RUN maintained a positive return in 2018, however, with both exceeding 80%. 

There was a significant difference in run time between the original and refactored code after we removed the nested loop and included a ticker index. The refactored code ran much faster, which would make it more suitable if Steve wants to analyze larger data sets than 12 stocks captured in 3013 rows.  

## Original Code Output & Run Time
![VBA_OriginalOutput_2017](https://user-images.githubusercontent.com/66224990/164052566-306322a7-ebcf-43e2-9ddc-39e39d4e1b56.png)<img width="253" alt="VBA_OriginalRunTime_2017" src="https://user-images.githubusercontent.com/66224990/164266380-4433bcc5-acee-45c1-81b2-a4089b76b10e.png">



![VBA_OriginalOutput_2018](https://user-images.githubusercontent.com/66224990/164052637-6a01b6e6-9fca-47cf-9c32-a36995639190.png)![VBA_OriginalRunTime_2018](https://user-images.githubusercontent.com/66224990/164266434-9e201a2e-6744-4b65-97f4-3e0d142aca63.png)



## Refactored Code Output & Run Time
<img width="324" alt="VBA_Output_2017" src="https://user-images.githubusercontent.com/66224990/164130299-d58c93cb-cab8-4dc9-8b49-da6d395ddf56.png"><img width="250" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/66224990/164130308-a320e12e-7306-4610-85b9-87f0b4adf09b.png">


<img width="325" alt="VBA_Output_2018" src="https://user-images.githubusercontent.com/66224990/164130320-566925f9-fc1c-4f88-916c-b24ca22c0b52.png"><img width="249" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/66224990/164130334-b3268da5-f973-4f93-af10-0d51538bac67.png">

# What Did We Change?

## Original Code Example
The original code below analyzed the data by looping through all 3013 rows for each ticker, confirming whether they matched the current ticker being analyzed. Since we had 12 tickers this means it needed to loop through 3013 rows 12 times. 

![VBA_Nested_Loop](https://user-images.githubusercontent.com/66224990/164266566-d3afd6c7-01f6-4c5d-9876-375ec35735da.png)

Although the macro got us the results we needed, the way the code was written was not very efficient and given that Steve wants to use this tool to analyze larger data sets it may result in very slow run times if the data set is much larger than 3013 rows.

## Refactored Code Example
The refactored code below removed the nested loop and used a ticker index, which enabled it to loop through all the rows once and gather the data for each ticker along the way. 

![VBA_TickerIndex](https://user-images.githubusercontent.com/66224990/164267347-abc84e12-0085-4c6f-954b-a9c934ba58f2.png)

This resulted in a much shorter run time since it removed the need to loop through all 3013 rows 12 times. With the inclusion of the 3 output arrays (tickerVolumes, tickerStartingPrices, tickerEndingPrices) we were also able to separate the tasks that were being performed. As the code looped through the rows it gathered the data for each ticker and then once all were completed it output that information.

## Summary: In a summary statement, address the following questions.

* What are the advantages or disadvantages of refactoring code?

* How do these pros and cons apply to refactoring the original VBA script?
