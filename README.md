# stock_analysis

## Overview of Project

The purpose of this analysis is to provide a tool to analyze stock returns. We began with a VBA script that analyzed specific years 2017 and 2018, which included data from 12 different stocks. We then refactored the code so we can run the analysis on data sets from any given year and potentially thousands of stocks. In order to make our script more efficient we removed the nested loop and created an index so that the code only needs to loop through the tickers one time. Using VBA's timer feature in the script we captured and compared the run times of the original and refactored code, results documented below. 


## Results: 

Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

## Original Code Output & Run Time
![VBA_OriginalOutput_2017](https://user-images.githubusercontent.com/66224990/164052566-306322a7-ebcf-43e2-9ddc-39e39d4e1b56.png)![VBA_Original_2017](https://user-images.githubusercontent.com/66224990/164052598-59dae291-6a9f-47c6-b573-52092f9a5ffe.png)


![VBA_OriginalOutput_2018](https://user-images.githubusercontent.com/66224990/164052637-6a01b6e6-9fca-47cf-9c32-a36995639190.png)![VBA_Original_2018](https://user-images.githubusercontent.com/66224990/164052675-754b45be-346d-4b66-b9e5-0d72d14a6c94.png)


## Refactored Code Output & Run Time
![VBA_Output_2017](https://user-images.githubusercontent.com/66224990/164048630-a5a4355b-c3ad-4ee7-93b4-717f0bff5cce.png)![VBA_Challenge_2017](https://user-images.githubusercontent.com/66224990/164048907-084f996c-3c58-4636-ab39-85bf29030445.png)


![VBA_Output_2018](https://user-images.githubusercontent.com/66224990/164048824-8f25a5f6-6243-426b-bf53-4ee6a3572b46.png)![VBA_Challenge_2018](https://user-images.githubusercontent.com/66224990/164048925-c310dd68-8127-41d1-9541-d97b96d0712c.png)



## Summary: In a summary statement, address the following questions.

* What are the advantages or disadvantages of refactoring code?

* How do these pros and cons apply to refactoring the original VBA script?