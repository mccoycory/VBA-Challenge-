# VBA-Challenge-
# Green Energy Stock Analysis 
# Overview 
The purpose of this project was to create a VBA script that would quickly allow the user to run analysis based on the percent change of a stock price and the total trading volume with in a given year. The data set given to us was a list of 12 green energy stocks and were were given the year 2017 and 2018.
# Results

![2017 Green Stock Return](https://github.com/mccoycory/VBA-Challenge-/blob/main/Resources/Stock%20Analysis%202017.PNG)

In 2017, 11 out of the 12 green stocks had a positive percent change. The stock that had a negative percent change in 2017 was ticker: TERP. Looking at this data set, it is easy to pick out that stock and further analysis could be done as to why this stock is performing poorly. 

![2018 Green Stock Return](https://github.com/mccoycory/VBA-Challenge-/blob/main/Resources/Stock%20Anaylsis%202018.PNG)

In 2018, there were 2 out of the 12 green stocks that reported postive percent change. These two stocks were ENPH and RUN. Likewise with 2017, we could look into ENPH and RUN and determine why they performed well year after year. I would invent my money in those two stocks. 

## VBA Scrpit 
### Orginal Code
Within the project were we able to run two different type of VBA code. This first type of code was almost hard code. The code was looping through each cell and each row and each column and output the results within the loop. This is a slower type code because the code is outputting the results for each stock ticker... 12 times. Please see below for code times. 

![2017 Original Code Time](https://github.com/mccoycory/VBA-Challenge-/blob/main/Resources/Orginal%20Code%202018.PNG)

![2018 Original Code Time](https://github.com/mccoycory/VBA-Challenge-/blob/main/Resources/Timer%20Original%202017.PNG)

### Refractor Code
During the refractor code, we were able to use the same concept of a for loop; however, we were able to establish an array for the output. This array would be outputted one time at the end of the code. Thus elminating the 12 line output loop. Please see below for code times

![2017 Refractor Code Time](https://github.com/mccoycory/VBA-Challenge-/blob/main/Resources/Timer%20Refractor%20Code%202017.PNG)

![2018 Reforactor Code Time](https://github.com/mccoycory/VBA-Challenge-/blob/main/Resources/Timer%20Refractor%20Code%202018.PNG)

# Summary 
The advantage to the refractor code is the single out put array. This will maximize your time if you are looking for the same set of stocks. However, if you are looking at many different stocks and want to perform the same analysis the orginal code would probably be better because you do not need to "hard code" the array in your VBA code. You can simply import the data and click run. 

I think it is always good to plot any stock against the S&P 500 and the specific industry sector return as well. This will help normalize the return on investment. Who knows, one might love green stocks but if the S&P 500 is giving the same amount of returns with less risk, I would put my money there. However, for these sets of stocks. I would invest in ENPH and RUN due to the year after year perforamce.
