# Stock Analysis

## Overview of Project

### Purpose
To refactor the existing Stock Analysis VBA script for greater sustainability when scaling for more exhaustive stock analysis. 

### Results

Using the data taken from both 2017 and 2018, we are able to see the total potential returns for each of the given stocks. We use the ***Returns*** column to indicate whether the overall investment had a potential financial gain or loss. The value for this column was calculated using the following:

```ticker_ending_prices/ticker_starting_prices-1``` 
where the mentioned variables were determined using a variation of the following ```if then``` logic:

```If Cells(j, 1).Value = ticker And Cells(j + n, 1).Value <> ticker Then```
where "```n```" is dependant on whether we are searching for the starting or ending price.

For easy visbility. the following code is used to format the background of the ***Return*** cell to either red or green, depending on whether it saw a loss or gain, respectively, for that year.

``` 
If Cells(i, 3) > 0 Then
'color the cell green
Cells(i, 3).Interior.Color = vbGreen
        
'color the cell red
ElseIf Cells(i, 3) < 0 Then
Cells(i, 3).Interior.Color = vbRed
```
    
As seen in Figure 1., 2017 was a strong year for many of the stocks that Steve has analyzed. If Steve desired, analyzing the impact of the market trends and milestones of 2017 would help him develop an understanding of a healthy market.

#### Figure 1. 2017 Analyses
![VBA_Challenge_2017](https://github.com/jdfiel/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

That said, based on Figure 2., it is clear that 2018 saw significant losses, especially in relation to 2017. If Steve desired minimum analyses of trends, investing in ENPH and RUN would be logical choices, as they maintained positive returns, contrary to the rest of the stocks.

#### Figure 2. 2018 Analyses
![VBA_Challenge_2018](https://github.com/jdfiel/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png).

Both figures also include the run-time of the code for both years. We can see that the refactored code has a near consistent run-time of ~0.125 seconds. This is a drastic improvement upon the previous run-time of ~0.80 seconds of the non-factored code. 

### Summary

#### Advantages and Disadvantages of Refactoring code

Obviously, the advantage of factoring code is to reduce the run time. It also has the benefit of cleaning up our code, removing any redundancies. The less code to decipher, the easier it is to update and/or debug. The largest disadvantage of refactoring code is that you are rewriting code. This opens up the possibility for creating new errors. Further, if the refactoring is not properly commented, it can result in confusion if someone else comes in to work with the code down the line - i.e. it's more work.

#### Pros and Cons specific to our VBA script
The most notable impact of the refactored code for our Data Analysis sheet is the number of times that VBA tells Excel to activate a different Worksheet. This also results in the collection and calculation of data and input of data into Excel itself occuring at a single point.

The logic for our unfactored code is as follows:
```
.Activate Data Sheet
Collect Data for Ticker i
.Activate Analysis Sheet
Input Data for Ticker i
.Activate Data
Collect Data for ticker i+1
[repeats n times]
```

For the non-factored coding logic, VBA is telling Excel to switch worksheets twice for every index in our Ticker array. The ```for loops``` and ```if then``` methods are also being initiated each time a ticker is being calculated as this code uses ```nested for loops```. Thus, the number of loops becomes (```n number of outer loops``` x ```n number of inner loops```), such that ```n``` equals the number of ```indexes``` in the ```ticker array```.

Compared to the factored code:
```
.Activate Data Sheet
Collect Data for Ticker i to n
.Activate Analysis Sheet
Input Data for Ticker i to n
```

As you can see, the refactored code needs to run only once, and the number of worksheets activated are drastically reduced to only two. The ```for loops``` and ```if then``` methods are not initiated exponentially like before, greatly reducing the run time. That said, the logic for the unfactored code is significantly easier to follow, especially since the concept and syntax of ```arrays```, in addition to their impact of ```for loops``` and ```if then``` methods can be difficult to learn.
