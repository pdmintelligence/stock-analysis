# Overview

The aim of the project is to use VBA script to analyze stock ticker outcomes associated with different business performance metrics such as total % returns on investment. There were 12 businesses and approximately 3K rows of data analyzed in this VBA script. The final goal was to look at each business listed and anaylze their total volume and return on an annualized basis using the periods of 2017 to 2018.  

# Results

The data shows that the businesses under review had significantly positive returns on investment for the period of 2017. Eleven (11) of the total twelve (12) firms under review experienced positive net returns on investment. The most successful investment as measured by % return on investment was DQ with a 199.4% return on investment. These trends were not replicated in the year 2018. In 2018, only two (2) of the assessed firms had positive returns on investment, ENPH and RUN. All other firms were negative for 2018, including DQ, which had a net negative % return on investment of -62.6%.



![VBA_Challenge_2017](https://user-images.githubusercontent.com/95975772/164000751-e52ee826-e9a7-4ceb-9a0e-dd806fef1392.png)




![VBA_Challenge_2018](https://user-images.githubusercontent.com/95975772/163997842-3f8abee0-1a1f-43c8-a096-a6de79bdf92d.png)


# Refactoring 

The process of refactoring code is an iterative process aimed at making code run faster and more efficiently. One of the key advantages of the refactored code is that it allows the user to input additional specifications for the year being reviewed. Additionally, the refactored code's usage of the index value tickerIndex coupled with for loops and the conversion of the stock ticker values from a set of string values into an array variable allowed the code to run more efficiently than it would have otherwise. This contrasts the original code which had the code run through all values in the excel sheet.


**Visual of time to execution in Refactored Code**

![image](https://user-images.githubusercontent.com/95975772/163999123-f6ea00dd-48b6-4be2-8dd3-5ccc7663c28d.png)




**Same visual but for original code**

![image](https://user-images.githubusercontent.com/95975772/163999371-5b8b105a-62a6-449a-8827-22f2fcf373df.png)


# Summary 

In this module we observed that the refactored code ran quicker due to modifications in the original code. The main advantage in the refactored code is that its usage of additional array variables allowed for the inclusion of more concise and efficient loops. This sped up the process of running the code by a significant amount of time. Whereas the original code took about 2 seconds to run, the refactored code executed in less than .2 seconds. 

There are no significant disadvantages of the refactored procedure; however, there is a higher learning curve when it comes to understanding and interacting with additional variable types and additional looping procedures. 
