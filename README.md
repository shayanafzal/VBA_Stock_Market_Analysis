# VBA-Challenge

## Overview of Project
### Purpose
The purpose of this task is to refactor the code so that it works efficiently with a larger data set of stocks.
### Background
Steve would like to do more stock market research for his parents. He would like to expand the dataset to include the entire stock market. In addition to that Steven would like to expand the number of years. This would result in significantly increasing the data set, making it important to have a code that runs faster.  

This will be accomplished by refactoring the code so that it loops through the data one time to collect the data.

## Results
### Stock Performance between 2017 and 2018

In the year 2017, 11 out 12 stocks had a positive rate of return.  
In the year 2018, only 2 out of 12 stock had a positive rate of return.  
Based on the above findings, 2017 was a good year for investing, where as in 2018, the markets took a downward turn that resulted in 10 out of 12 stock giving a negative return.  




### Execution Times of the Original and Refactored Script
The refactored script run much faster than the original. 

When running the analysis for 2017, the refactored code had a runtime of 0.17 seconds, compared to 0.87 seconds for the original code. Please refer to the screen shots below.

#### 2017 - Original Code
![Orgignal Code 2017](https://github.com/shayanafzal/VBA-Challenge/blob/b005d143354040bbba311a69c4648f2ad32905dc/Resources/VBA_Challenge_2017.png)

#### 2017 - Refactored Code
![Refactored Code 2017](https://github.com/shayanafzal/VBA-Challenge/blob/b005d143354040bbba311a69c4648f2ad32905dc/Resources/VBA_Challenge_2017_Refactored.png)


When running the analysis for 2018, the refactored code had a runtime of 0.16 seconds, compared to 0.86 seconds for the original code. Please refer to the screen shots below.


#### 2018 - Original Code
![Orignal Code 2018](https://github.com/shayanafzal/VBA-Challenge/blob/b005d143354040bbba311a69c4648f2ad32905dc/Resources/VBA_Challenge_2018.png)

#### 2018 - Refactored Code
![Refactored Cdoe 2018](https://github.com/shayanafzal/VBA-Challenge/blob/b005d143354040bbba311a69c4648f2ad32905dc/Resources/VBA_Challenge_2018_Refactored.png)


### Screen shots and Code
Below is the portion of the code that has been refactored. 
![Refactored Code](https://github.com/shayanafzal/VBA-Challenge/blob/0473229ccd1c6351043b7e45906b6826334cb897/Resources/Code%20Refactored.png)

Looking at the refactored code above it can be noted that the refactored code loops through the data once and collects all the information. Since it is looping though the data only once, it run about 5 to 6 times faster than the original code. 

## Summary
### Advantages or Disadvantages of Refactoring Code
#### Advantages
* Refactored code takes fewer steps, there by decreasing the amount of time required to run the analysis
* The code uses less memory and computing power
* Refactoring the code can make it easier for others to read the code and follow the logic behind it.
#### Disadvantages
* Refactoring the code can create errors that will require more time to fix
* It may not be to a businesses financial advantage to invest money into refactoring code
* The original code is working, and refactoring it could cause disruptions

### How do the pros and cons apply to refactoring the original VBA Script
The refactored code runs approximately 5 to 6 times faster then the original code. We are working with a small data set in this exercise, however, the refactored the speed up the process 5 to 6 times will be noticeable when working with a large data set.  

When Steve starts using this notebook to conduct an analysis of the complete stock market over a large time span, this improvement in time will be noticeable.  

It took sometime to improve the code, however, the code is now more efficient, concise and easy to follow. Hence it was well worth the time to refactor this code making it better for Steve.

The excel document with the Stock Analysis code can be accessed
[here](https://github.com/shayanafzal/VBA-Challenge/blob/8e0dc23fc8120bd4d31099274ac275ba37aa4813/VBA_Challenge.xlsm).




