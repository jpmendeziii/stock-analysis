# Stock Analysis with VBA and Excel
## Overview - VBA project
### Purpose 
The purpose of this project and subsequent analysis is to edit, refactor a stock market dataset with VBA solution code to loop through all the data one time to collect an entire dataset.  Refactoring is a key part of the coding process – when refactoring code, adding functionality is not the key – the goal is to gain efficiencies in concise coding or reducing memory / resource utilization.  Therefore, part of the deliverable is to also verify the changes made through refactoring of the VBA script were successful in making the code more efficient.  Please refer to the [VBA_Challenge.xlms](VBA_Challenge.xlsm) for reference - this spreadsheet contains all relevant material.  The 'All Stocks Analysis' sheet contains all buttons shown here.  
![2017 and 2018 VBA Buttons](2017_and_2018_VBA_Buttons.png)

## Results
### Deliverable Requirements and Coding Examples
![Question_1_Code_Snip](Question_1_Code_Snip.png)

![Question_2_Code_Snip](Question_2_Code_Snip.png)

![Question_3_Code_Snip](Question_3_Code_Snip.png)

![Question_4_Code_Snip](Question_4_Code_Snip.png)

4. The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
A loop was created that will loop over all the rows in the spreadsheet. Inside the loop, a script was created that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.

### Stock Performance and Refactor Time Results
![2017_Original_VBA_code.png](2017_Original_VBA_code.png)
![2017_REFACTORED_VBA_Code.png](2017_REFACTORED_VBA_Code.png)

![2018_Original_VBA_Code.png](2018_Original_VBA_Code.png)
![2018_REFACTORED_VBA_Code.png](2018_REFACTORED_VBA_Code.png)



## Summary and Conclusions

Refactoring helps code understanding.  Refactoring encourages each developer to think about and understand design decisions, in the context of collective ownership / collective code ownership.  After refactoring, the code is fresher, easier to understand or read, less complex and easier to maintain. Disadvantages of Code Refactoring: Time Consuming: You may have no idea how much time it may take to complete the process. It may also land you into a situation where you have no idea where to go.  

