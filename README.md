# VBA Challenge Analysis

## Stocks Overview
Steve's parents want to invest money with him helping guide the way. He wants to give them the best advice possible based on the data he has access to. Rather than using traditional excel funtions, we cut through to show Steve what he really wants to see with just a couple lines of code using macros. 

### The Macro in Depth
The code creates a more complex pivot table that displays values based on user inputs (as long as data is availalbe) and conditional formatting based on the rate of return for each stock. The InputBox function is one of the first pieces of code in the analysis. 

<img width="481" alt="Screen Shot 2021-06-22 at 10 19 50 PM" src="https://user-images.githubusercontent.com/82982180/123030538-930cfb00-d3a8-11eb-830e-209008b6b1ac.png">

It pops up similarly to a message box but in this case it allows the user to choose the year of data they wish to see.

<img width="481" alt="Screen Shot 2021-06-22 at 10 21 06 PM" src="https://user-images.githubusercontent.com/82982180/123030612-ae780600-d3a8-11eb-9279-324cfa0bef66.png">

In this use case there are only two options to choose from, either 2017 or 2018. Once the user chooses the year then the code is ran and the sheet represents data from the chosen year. This is done so by the yearValue variable which is chosen by the user's input and then used throughout the rest of the code to ensure that the correct data is displayed. 

<img width="389" alt="Screen Shot 2021-06-22 at 10 30 04 PM" src="https://user-images.githubusercontent.com/82982180/123035137-6fe64980-d3b0-11eb-9cc4-147a501cbc00.png">
<img width="381" alt="Screen Shot 2021-06-22 at 11 21 35 PM" src="https://user-images.githubusercontent.com/82982180/123035232-a0c67e80-d3b0-11eb-8fc5-614cca9ed987.png">

I used for loops and if statements to have the proper values in the corresponding cells. 

<img width="529" alt="Screen Shot 2021-06-22 at 11 29 52 PM" src="https://user-images.githubusercontent.com/82982180/123035980-e6d01200-d3b1-11eb-9cb6-7117c8d78c46.png">

## Results
### 2017 Stock Analysis
In 2017 the only stock that had a loss of a return was "TERP". All other stocks had a positive return. The stocks with the greatest return in 2017 include DQ (199% return), SEDG (185% return), ENPH (130% return), and FSLR (101% return). 

<img width="355" alt="Screen Shot 2021-06-22 at 11 49 27 PM" src="https://user-images.githubusercontent.com/82982180/123037596-842c4580-d3b4-11eb-975e-05c9ba85d8bd.png">

### 2018 Stock Analysis
In 2018 a completely different story is told where there are only two stocks that had a positive return. Those two stocks were ENPH (82% return) and RUN (84% return). 

<img width="355" alt="Screen Shot 2021-06-22 at 11 48 46 PM" src="https://user-images.githubusercontent.com/82982180/123037511-67900d80-d3b4-11eb-950c-84f8552c89ab.png">

### Overall Stock Analysis
Both ENPH and RUN stocks had postive returns in 2017 and 2018. Only ENPH had an outstanding rate in 2017 and was barely beat out by RUN with a slightly better return rate in 2018. Based on ENPH and its consistent performance across both years, it would be my front-runner to invest in. RUN would be my second recommendation, seeing as it was the only other stock that had a positive return rate in 2018 and it's improvement from 2017 to 2018 might be a sign of a trend. It also might be worth while to be on the lookout for SEDG which had an off year in 2018 with a -8% return rate but it might be a good time to buy in as it had the second best return rate in 2017. 

### 2017 Run Time Comparison
As you can see below, the refactored code made the run time over 8 times faster than the original code. 
<img width="252" alt="2017_Normal" src="https://user-images.githubusercontent.com/82982180/123039648-26016180-d3b8-11eb-9dc7-c71bdc152fcf.png">
<img width="252" alt="2017_Refactored" src="https://user-images.githubusercontent.com/82982180/123039659-2a2d7f00-d3b8-11eb-8107-43ac148988f5.png">

### 2018 Run Time Comparison
As you can see below, the refactored code made the run time over 8 and a half times faster than the original code.
<img width="252" alt="2018_Normal" src="https://user-images.githubusercontent.com/82982180/123039736-49c4a780-d3b8-11eb-9e2d-58e6b7280e22.png">
<img width="252" alt="2018_Refactored" src="https://user-images.githubusercontent.com/82982180/123039746-4e895b80-d3b8-11eb-98ab-799940e0e0d9.png">

## Advantages of Refactoring Code
An overall faster run time when running the code later seems to be the main advantage. It can help the code become more efficient. The efficiency doesn't only apply to the time it takes to run the code, but possibly the understanding of the code. When you refactor the code you have to fully understand what you have written and also have a better way to write it that improves the code. This helps you better understand the ins and outs while giving you an opporunity to improve your past work. 

## Disadvantages of Refactoring Code
It is a time consuming process and if the dataset is not large then there isn't much added benefit. In terms of this scenario, the code run time went from being around a half a second to a seven hundredth of a second. In my experience, this doesn't help me gain much. If the code took 30 minutes to run and then the refactor code took 7 seconds to run, then I could see the benefit. That is essentially the difference that the refactored code would have had on a much larger dataset. 
Refactoring the code does give another opportunity to mess up what you already have in place. If it is something that's worth doing, I believe it is worth while to spell out those steps before taking them. The given code for the challenge helped direct me to the right places and gave me an order in which I needed to implement the new code. Without the established steps in place it would be easy to edit the wrong pieces of code. 
