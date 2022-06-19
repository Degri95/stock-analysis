# Stock-Analysis with VBA

## Overview

In this project we've helped our friend Steve by writing a VBA script to help filter and analyze some green energy stock data from 2017 and 2018 to present to his parents. After the initial script was written, our challenge was to see if refactoring our code could make it run more efficient.

## Results

### initial code & analysis

Our initial code used nested [for loops](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/fornext-statement) and [conditional statements](https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/using-ifthenelse-statements) to iterate through our data.

![original code](/Resources/original_code.PNG)

 Once the conditions were met the data was stored in a variable and output in our analysis sheet at the end of the loop. We successfully were able to export our data and format it to make our analysis. We also created buttons to run the script give ease of access to our friend Steve since he will be using the excel sheet as the end user.

![buttons](/Resources/buttons.PNG)

### Stock data

![2017 stocks](/Resources/Stocks_2017.PNG) ![2018 stocks](/Resources/Stocks_2018.PNG)

In our data we see that our stocks were in a positive trend in 2017 with the exception of TERP. In 2018 only 2 stocks remained in the green in a seemingly bear market in the industry (ENPH and RUN).

### Refactoring 

Our initial code did the job but could it do it more efficiently? In the refactored code we've removed the nested loop and added arrays to store our data instead of updating the same variables. We had a variable outside our loop working as our ticker as we iterated through our data. 

![updated code](/Resources/refactored.PNG)

After our data was stored in our arrays we used another for loop to ouput it into our analysis sheet.

![updated output code](/Resources/refactored2.PNG)

### Comparison

I had timers set in both of our scripts to see how long they would take to run. Below is the orginal scripts excecution time

![original script 2017](/Resources/Old_2017.PNG) ![original script 2018](/Resources/Old_2018.PNG)



And now our refactored codes excecution time

![Refactored script 2017](/Resources/VBA_Challenge_2017.PNG) ![Refactored script 2018](/Resources/VBA_Challenge_2018.PNG)


Looking at the comparison we see the the refactored scripts time is **substantially** faster. The code now excecutes 7 times faster than the original.


## Summary

### Advantages of refactoring code:

- Refactoring code can significantly speed up your code and reduce memory usage.
- A fresh pair of eyes or another perspective on code can be very helpful.
- Finding unnecessary code and better algorithms can be done through refactoring. 

In refactoring the VBA code we've sped up our excecution time and made it more readable. 

### Disadvantages

- Refactoring code can lead to bugs and errors if not precise on changes being made
- It can be very time consuming and stressful
- Refactoring may be rushed if there's a deadline.

The biggest disadvantage to refactoring the VBA code would be the potential for errors. All though Through testing I was able to confirm that all of the changes were successful.









