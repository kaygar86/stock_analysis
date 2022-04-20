# stock_analysis

## Overview of Project

The purpose of this analysis is to provide a tool to analyze stock returns. We began with a VBA script that analyzed specific years 2017 and 2018, which included data from 12 different stocks. We then refactored the code so we could run the analysis on data sets from any given year and potentially thousands of stocks. In order to make our script more efficient we removed the nested loop and created an index so that the code only needs to loop through the tickers one time. Using VBA's timer feature in the script we captured and compared the run times of the original and refactored code, results documented below. 


## Results: 

There was no change in the analysis result of the stocks when we ran the original code and the refactored code. In 2017 the majority of the stocks we analyzed showed a positive return, with a few showing a return exceeding 100%. The analysis of 2018 showed the majority of stocks with negative returns. ENPH and RUN maintained a positive return in 2018, however, with both exceeding 80%. 

There was a significant difference in run time between the original and refactored code after we removed the nested loop and included a ticker index. 

## Original Code Output & Run Time
![VBA_OriginalOutput_2017](https://user-images.githubusercontent.com/66224990/164052566-306322a7-ebcf-43e2-9ddc-39e39d4e1b56.png)![VBA_Original_2017](https://user-images.githubusercontent.com/66224990/164052598-59dae291-6a9f-47c6-b573-52092f9a5ffe.png)


![VBA_OriginalOutput_2018](https://user-images.githubusercontent.com/66224990/164052637-6a01b6e6-9fca-47cf-9c32-a36995639190.png)![VBA_Original_2018](https://user-images.githubusercontent.com/66224990/164052675-754b45be-346d-4b66-b9e5-0d72d14a6c94.png)


## Refactored Code Output & Run Time
<img width="324" alt="VBA_Output_2017" src="https://user-images.githubusercontent.com/66224990/164130299-d58c93cb-cab8-4dc9-8b49-da6d395ddf56.png"><img width="250" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/66224990/164130308-a320e12e-7306-4610-85b9-87f0b4adf09b.png">


<img width="325" alt="VBA_Output_2018" src="https://user-images.githubusercontent.com/66224990/164130320-566925f9-fc1c-4f88-916c-b24ca22c0b52.png"><img width="249" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/66224990/164130334-b3268da5-f973-4f93-af10-0d51538bac67.png">




## Summary: In a summary statement, address the following questions.

* What are the advantages or disadvantages of refactoring code?

* How do these pros and cons apply to refactoring the original VBA script?
