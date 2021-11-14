# Stock Analysis

## Overview: Stock Analysis using VBA Coding

### Purpose:
 >To select year with assigned stock tickers list, and return analysis table columns for `Ticker, Total Daily Volume, and Return` with click of a button.  

### Results:

 - #### VBA Coding:
 
 >Ticker array as a string (text) with indexing numbers for which tickers will be analyzed:
 >
 >![Dim_Tickers_as_string](https://user-images.githubusercontent.com/92836648/141665889-d99867ce-c6c7-4c73-9583-025720dc44f0.png)
 
 >Row Count coding to make sure every row of data is evaluated:
 >
 >![Data_row_counts](https://user-images.githubusercontent.com/92836648/141665888-7dd27be9-c688-4c7c-9abb-a62bde745be4.png)

 >Arrays to store data counts for output results:
 >
 >![Refractored_arrays_code](https://user-images.githubusercontent.com/92836648/141664956-cad2fb98-03a0-45e0-a5e0-e5d83377891f.png)

 >Outer and Inner `For Loops` for evaluation data set:
 >
 > - Outer Loop ( i ) is used to identify which stock ticker from Ticker array is being evaluated
 >
 > - Inner Loop ( j ) is the evaluation statement for totaling stored results back to Output arrays
 >
 >![Inner_outer_loops](https://user-images.githubusercontent.com/92836648/141665890-f191030d-31ba-4dd5-9ea1-f3f7f1611971.png)

 >Conditional Formatting rules for evaulated results:
 >
 >![background_formatting_value_rules](https://user-images.githubusercontent.com/92836648/141665885-228e55d9-03a0-4082-bcc4-b653fcd63337.png)

 - #### Comparison Results of Stock Analysis:
 >Analysis of 2017 and 2018 data sets indicates two stocks maintained positive returns:
 >
 > |      2017     |      2018     |    
 > | ------------- | ------------- |
 > | ENPH +129.5%  | ENPH +81.9%   |
 > | RUN    +5.5%  | RUN  +84.0%   |
 > 
 >![VBA_Challenge_2017](https://user-images.githubusercontent.com/92836648/141664529-f4f67063-3ae9-46ed-b640-d589ab664290.png) 
 >
 >![VBA_Challenge_2018](https://user-images.githubusercontent.com/92836648/141664532-95385731-07bb-434f-8f6c-3213d60a2948.png)

 - #### Comparison Results of Run Time Codes:
 >VBA coding performance evaluation runtimes between standard & improved coding:
 >
 > |                          |      2017     |     2018      |
 > | ------------------------ | ------------- | ------------- |
 > | Non-refactor (standard) | 0.93 seconds  | 0.87 seconds  |
 > | Refactor (improved)     | 0.87 seconds  | 0.83 seconds  |
 >
 >Non-refactor (standard) coding runtime results:
 >
 >![Green_Stock_2017](https://user-images.githubusercontent.com/92836648/141664506-13693a8b-26c2-4c73-8fe6-c83cd5e5b159.png)
 >![Green_Stock_2018](https://user-images.githubusercontent.com/92836648/141664527-87de2974-8da7-4a38-be54-3204122a01a2.png)
 >
 > Refactor (improved) coding runtime results:
 >
 >![VBA_Challenge_2017](https://user-images.githubusercontent.com/92836648/141664529-f4f67063-3ae9-46ed-b640-d589ab664290.png) 
 >![VBA_Challenge_2018](https://user-images.githubusercontent.com/92836648/141664532-95385731-07bb-434f-8f6c-3213d60a2948.png)

### Summary:

 - #### Advantages vs Disadvantages of Refactoring Code:
 >
 > - ##### Advantage concluded from the coding is the Output arrays and the data type selected for data storage capacities.
 >
 >![Refractored_arrays_code](https://user-images.githubusercontent.com/92836648/141664956-cad2fb98-03a0-45e0-a5e0-e5d83377891f.png)
 >
 > - ##### Disadvantage concluded for the coding is Ticker arrays being limited to only coded index without the ability update unless VBA coding is a skillset for end user
 >
 >![Dim_Tickers_as_string](https://user-images.githubusercontent.com/92836648/141665889-d99867ce-c6c7-4c73-9583-025720dc44f0.png)

 - #### Advantages vs Disadvanteges between Non-refactored and Refactored Code:
 >
 > - ##### Advantage concluded from the coding is the Output arrays and the data type selected for data storage capacities, and along with Refactored coding adding `Dim tickerVolume`.
 >
 >Non-refactored:
 >
 >![Non_refractored_arrays_code](https://user-images.githubusercontent.com/92836648/141664954-ab63d18d-5f8a-43bf-870e-9dc1c9d8c6e4.png)
 >
 >Refactored:
 >
 >![Refractored_arrays_code](https://user-images.githubusercontent.com/92836648/141664956-cad2fb98-03a0-45e0-a5e0-e5d83377891f.png)
 >
 > - ##### Disadvantage concluded for the Non-refactored and Refactored coding is Ticker arrays being limited to only coded index without the ability update unless VBA coding is a skillset for end user
 >
 >![Dim_Tickers_as_string](https://user-images.githubusercontent.com/92836648/141665889-d99867ce-c6c7-4c73-9583-025720dc44f0.png)
