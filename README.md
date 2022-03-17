# Analysis of DQ Stocks
**Overview**

**Purpose**
The client is researching data on green energy for his parents.  He provided us with an Excel file of stock data to analyze. For each stock, we determined total daily volume and yearly return, as well as focused heavily on the stock, DQ.  We then refactored our code, making it more efficient so that it could handle analyzing thousands of stocks instead of a few at a time.

**Results**

**2017**
For the 2017 stocks category, all but one stock had a positive percentage of return, with only TERP bringing in a negative percentage.  
![VBA_Challenge_2017](https://user-images.githubusercontent.com/100445222/158898191-fc4411f8-2077-4a81-88a7-a5b76b4810e8.png)

**2018**
The 2018 stocks category had a greater number of negative returns, with only two positive returns for ENPH and RUN.  
![VBA_Challenge_2018](https://user-images.githubusercontent.com/100445222/158898200-c5a198b6-4fb4-44ce-8389-4da530d8b374.png)

**2017 vs 2018**
The original VBA script applied only to 2018 data, whereas the refactored script combined both 2017 and 2018. While we could copy the code to also create the same model for the 2017 data, combining both in the refactored script creates a less chaotic workspace and helps better organize the data in one worksheet.  

**Original vs refactored script**
Because the refactored code is combining all aspects into one Macro, we can time how long it takes the program to execute the code.  For the original script, however, because each piece is a separate Macro, we are unable to time the entire set of code.  This must be done individually, which is more time consuming.  

**Summary**

**Advantages**
Advantages to refactoring code include editing out unnecessary coding language or reorganizing code to create a clean, concise workspace.  An original script may be very long and chaotic, and refactoring can make it more easily understandable to the person looking at it.  

**Disadvantages**
Disadvantages to refactoring code include broken code language.  If verbiage is rewritten or removed, it may not execute the same action as the original script.  

**How do these pros and cons apply to refactoring the original VBA script?**
The original script was very long and split into separate Macros.  With refactoring, the script became more organized and combined so that it could be ran as a single Macro.  Fortunately, the cons did not apply to this particular demonstration as there was no broken code language to factor in.
