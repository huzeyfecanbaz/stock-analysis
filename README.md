# stock-analysis

## Overview of Project
This project mainly serves to reduce the code written throughout the Module 2. Refactor is using by computer programmers to tweak the codes they have wrote for particular tasks and make it more meaningful by implementing smaller line codes and expressions.
Module 2 was serving to teach the Virtuall Basic for Applications(VBA)in Excel. The project started with a data set and lead the programmer to organize various codes and apply them in different sheets such as DQ Analysis, All Stock Analysis etc. Additionally, it let the programmers to deep dive by using VBA and improve their skills not only in Excel but also by using some online applications and websites such as effective googling, stackoverflow, and so on.  
The Module set criteria and let the programmer to search new things and apply them in order to find an accurate solution to the tasks in which different ways that asks programmer to implement a hard-code directly, or integrate the code to a button so that the user do not need to worry about anything but just click a button and VBA will take care of the rest for the user.

## Results
As it is explained in Background of the Challenge, Refactoring is a key part of the coding process where the programmers are making small touches to reduce the lines of the codes and make them more meaningful and more faster in the progress. 
In order to complete the refactoring, I have started to define the terms used in the code and their data type by using Dim. For instance, the start time, end time, year value and tickerindex is defined as follows;
    Dim startTime As Single
    Dim endTime  As Single
    Dim yearValue As String
	Dim tickerIndex As Byte
Use of startime and endtime let the programmer to compare the progress time before and after the refactor, starttime and endtime functions are added to the code and screenshots are saved and recorded in GitHub and explained in this written statement. 
The code is basically started by asking a question to the user:"What year would you like to run the analysis on?" by using
	yearValue = InputBox("What year would you like to run the analysis on?")
Then, the code activated the worksheet the analysis are going to report;
	Worksheets("All Stock Analysis").Activate
After this step, the arrays are created and loops initialized and code finished with endtime function.
The data types used in the code were single, and string before starting the refactoring. As it is reported these data types made the code's run time 0.2656 seconds.
![AllStockAnalysis Without Refactoring(2017)]([Report/AllStockAnalysis_without_refactoring(2017).png](https://github.com/huzeyfecanbaz/stock-analysis/blob/daa5b9a4d33696a99393069fbf12e8b88af4e26b/Resources/AllStockAnalysis_without_refactoring(2017).png)
![AllStockAnalysis Without Refactoring(2018)]([Report/AllStockAnalysis_without_refactoring(2018).png](https://github.com/huzeyfecanbaz/stock-analysis/blob/daa5b9a4d33696a99393069fbf12e8b88af4e26b/Resources/AllStockAnalysis_without_refactoring(2018).png)

During the refactoring the index data type used which let me to reduce the size of the data and make it run faster. The results shown below for both 2017 and 2018 Analysis;
![VBA Challenge 2017]([VBA_Challenge_2017.png](https://github.com/huzeyfecanbaz/stock-analysis/blob/daa5b9a4d33696a99393069fbf12e8b88af4e26b/Resources/AllStockAnalysis_without_refactoring(2017).png)
![VBA Challenge 2018]([VBA_Challenge_2018.png](https://github.com/huzeyfecanbaz/stock-analysis/blob/daa5b9a4d33696a99393069fbf12e8b88af4e26b/Resources/VBA_Challenge_2018.png)
The results explicitly exhibited that the refactoring works fine by reducing the code running time.

## Summary
As a summary, it has been seen that the refactoring does not mean to create something new in the existing code but instead tweak that code in a way that reduces the data type and does some looping changes so that the code run more faster. 

- What are the advantages or disadvantages of refactoring code?
Refactoring is a programming that allows the computer programmers to make long line codes shorter and more faster in the progress. These are the most definite advantages of the refactoring. Additionally, it lets the programmers to have a better quality of the code with less progress time. For instance, the code I was running before the refactoring was taking 0.2656 seconds for the All Stock Analysis for 2017, however, it took only 0.07 seconds after refactoring. The same thing happened for 2018 data, the code was ran in 0.2656 seconds before the refactoring and it took only 0/085 seconds after rewriting the code with some small changes. Although I could not find any disadvantages, I may say that it was hard to refactor my own code and I cannot foresee how long it will take to refactor for someone else's code.


- How do these pros and cons apply to refactoring the original VBA script?
Throughout the challenge, I was writing long codes to express every single thing and it was making coding time and accuracy longer. I have put startime and endtime function to see the codes' run time. After implementing the refactoring, the results exhibited that the refactoring made the code run faster and make the coding easier since I reduced the number of lines. As it is mentioned in the previous question, there was no disadvantages seen in the refactoring process. The only cons was that the refactoring is a challenging process and it takes a while to understand the codes system in the first place and it is a fact that the refactoring is starting after understanding a running code and the elements' function in the code as a whole.

