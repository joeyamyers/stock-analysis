# stock-analysis

## Overview
### Purpose:
The original goal of undertaking this analysis is to help Steve's parents make an informed investment decision regarding tweleve different green energy stocks. I used stock performance data from 2017 and 2018 to accomplish the analysis and provided Steve a workbook to analyze a dataset at a click of a button. After successfully performing this analysis for Steve, he wants to expand the dataset and to refactor the original code in order to improve the performance for future stock analyses. Refactoring the code will not change the contents of the code, but only improve its performance.

## Results
### Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
<img width="358" alt="StockAnalysis2017" src="https://user-images.githubusercontent.com/96351306/149578754-83160851-5159-4e2b-8156-a997de91a884.png">
<img width="353" alt="StockAnalysis2018" src="https://user-images.githubusercontent.com/96351306/149578765-2764e9bb-d156-4b1f-b1b4-d849481ad2dd.png">

Stock returns in 2017 were generally positive as all but "TERP" produced positive returns. Conversly in 2018, the selected green energy stocks mostly had a down year, as only "ENPH" and "RUN" generated positive returns. The stock Steve's parents selected, "DQ", outperformed the selected market in 2017 generating an outstanding return of 199.4%, but produced negative returns in 2018. They should be cautious in investing in "DQ" in the future.

### Original script vs Refactored Script:
![Original_script_17](https://user-images.githubusercontent.com/96351306/149578797-6bc17976-662f-4453-b773-94a5100d4784.png)
![Original_script_18](https://user-images.githubusercontent.com/96351306/149578817-0a1a7939-7532-43e0-b6b2-0b5e835c5b20.png)

![VBA_Challenge_2017](https://user-images.githubusercontent.com/96351306/149578856-2e249b7e-2c50-4cc2-8239-4186a5b186a9.png)
<img width="246" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/96351306/149578879-e724618f-3da7-4d55-b5ce-afafe60075d2.png">

From the images above, it is apparent that the refactoring process was a success. The refactored script in both years ran was executed quicker than the original script, while accomplishing the same stock analysis.


## Summary: In a summary statement, address the following questions.
### Advantages or disadvantages of refactoring code?
Refactoring code can be likened to proofreading an essay to make it more clear and concise. It does not change the funcitionality or purpose of the original script, however the advantages come from making the code more efficient. Making code more efficient can mean several things like making the code easier to read for other users through improving the logic of the code and taking fewer steps and also by using less memory.

Refactoring can become a disadvantage if the integrity of the code is not maintained throughout the process. Refactoring, like in this instance, can be quite time consuming and may not be worth the investment of time.

### How do these pros and cons apply to refactoring the original VBA script?
In this case, the script is now more efficient and can handle larger datasets, even though it took more steps to accomplish the refactor. However, the refactoring process was not a disadvatage as the refactor was well worth the investment of time and the integrity of the code was not altered.
