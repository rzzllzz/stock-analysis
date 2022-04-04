# Stock Analysis with VBA

## Overview of Project
### Purpose
  The purpose of this project was to refactor Microsoft Excel VBA code to be adaptable for use on the data already provided as well as other datasets. The code created for this workbook was made with the intention to analyze past datasets to inform future decisions. 
  
### Background
  The data provided includes information for various stocks as well their performance, such as the ticker, opening and closing price, volume, etc. The main goal for analyzing this data set was to find the total daily volume as well as the percentage return on each stock in an accessible table with colors to indicate performance at a glance. 

## Results
### Analysis
  As stated, the code used in the challenge was refactored from previously created code that only looked at two datasets. The goal after refactoring was to create code that would run much faster and be applicable to other datasets. Thus, the best place to start was by copying the old code. The starter code provides an outline to follow, as seen below. Following these steps while keeping the flow of the original code in mind led to a successful outcome. The concepts of arrays and assigning variables and creating loops still applied to the refactored code, allowing for more focus on applying the code beyond the two datasets.   
 
![screenshot_AllStocksAnalysisRefactored_using_outline](https://user-images.githubusercontent.com/101225282/161493378-b7199d31-cf80-49d7-9399-da7d5d8ba1ef.png)

  When it comes to the results of the analysis of the stocks over the course of 2017 and 2018, it is easy to see that 2017 was a much better year for returns on each stock. ENPH and RUN were the only stocks to have a positive return percentage across both years. While this might not be enough information to make a sound investment decision, Steve can use the code we created to run analyses for previous years.
 
![screenshot_of_2017_all_stocks_analysis_refactored](https://user-images.githubusercontent.com/101225282/161495166-398c3445-9046-4958-90f5-e51143d300c5.png)

![screenshot_of_2018_all_stocks_analysis_refactored](https://user-images.githubusercontent.com/101225282/161495168-ddc115e6-278d-4c06-9b08-7f2bb10086df.png)

## Summary
### Advantages vs. Disadvantages of Refactoring Code
  One advantage to refactoring code is that it saves time. The older code served as a sort of "outline" for the new code. A disadvantage to refactoring is that it might feel a little cluttered while you are scrolling through the lines to edit the code as needed. However, the old code was already meticulously organized with consistent indents and line spacing. While it might seem overwhelming to scroll through the code to find the areas that need updating, there is no need to waste time focusing on formatting each line since it had already been done.

## Advantages vs. Disadvantages of Refactoring the Original VBA script 
  The original macro Sub AllStocksAnalysis() had a run time of .86 seconds for 2018 and .87 seconds for 2017. In contrast, the macro for Sub AllStocksAnalysisRefactored() had a run time of 0.13 seconds for 2018 and 0.13 seconds for 2017. Refactoring benefitted the code's performance by improving runtime.
  
![All_Stocks_Analysis_2018_run_time](https://user-images.githubusercontent.com/101225282/161492190-e4e124ac-12a5-4ef9-828b-57b7c935151c.png)
![All_Stocks_Analysis_2017_run_time](https://user-images.githubusercontent.com/101225282/161492197-51b5080d-ef67-42c5-bcc6-571d7673b884.png)

![Run_Time_for_Refactored_Stock_Analysis_for_2018](https://user-images.githubusercontent.com/101225282/161492422-20ae46d8-6e3d-439b-a932-e92bc5dbeb28.png)
![Run_Time_for_Refactored_Stock_Analysis_for_2017](https://user-images.githubusercontent.com/101225282/161492439-684bfbc4-c67e-43f3-9aea-716ff30fdade.png)
