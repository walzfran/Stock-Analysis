# Stock-Analysis


Module Two Challenge

# VBA of Wall Street in Excel

## Overview of Project

### This project is to assist our friend, Steve, a new graduate in finance, compare data gathered from multiple alternative energy companies. He has a new clinet, his parents, and they have already invested their money in _DAQO New Energy Corp_. Steve wants a way for him to better understand what companies in the alternative energy industry are strong investments so he can yeild a successful return for his parents. 

## Results

### When looking at the outcomes from the two different years of data analysis it appears that DAQO New Energy Corp - DQ, was successful and had a very high return in 2017 of almost 200%. The image below shows all of the tickers and their returns. 
![2017_Returns](https://github.com/walzfran/Stock-Analysis/blob/main/2017_Stock_Returns.png)

* If you were to only go off of the data from 2017, it would appear as though Steve's parents made a great choice investing in DQ as it yeilded the highest return compared to the other alternative energy companies that we are looking at.
* There are a few other companies that did very well in 2017 those being ENPH, SEDG, and FSLR akk with returns over 100%. Only one company is at a loss and that is TERP but there are a couple other compaies with very low returns such as AY and RUN, both having less than 10%. 

### Now, looking at the outcomes from 2018, DQ has seen a negative return as depicted in the image below. 
![2018_Returns](https://github.com/walzfran/Stock-Analysis/blob/main/2018_Stock_Returns.png)

* We can snow see that DQ is showing a negative return of over 60%, but most of the companies are showing a loss. Only 2 of the companies experienced a positive return - ENPH and RUN. 
* In 2017, RUN had a very low return of only 5.5% but now it has the highest return of 84%. With it being positive returns both years in a row, it may be a great option for Steve to invest his parents portfolio into. 
* Another great option would be ENPH which has had good returns both years, in 2017 and 2018 it was one of the highest yeilds with 2017 being 129.5% and 2018 being 81.9%.
* If it were me and I was looking to invest after seeing this information, I would put my money in ENPH as it has been very consistent and at the top of yeilds both years. 

### I was able to successfully run the analysis macros in both the refactored and non refactored coding, but the results were yeilded in very different times - below are the images from the message box that displays how quickly for 2017 with the first image being from the inital macro and the second image being from the refactored macro. 

![2017_Timing](https://github.com/walzfran/Stock-Analysis/blob/main/2017_Stocks_Analysis.png)
![2017_Timing_Refactored](https://github.com/walzfran/Stock-Analysis/blob/main/VBA_Challenge_2017.png)

### Below are the same screenshots as above but for 2018, with the first image being the initial macro and the second being the refactored.

![2018_Timing](https://github.com/walzfran/Stock-Analysis/blob/main/2018_Stocks_Analysis.png)
![2018_Timing_Refactored](https://github.com/walzfran/Stock-Analysis/blob/main/VBA_Challenge_2018.png)

### Because I am new to coding, I had to start at 0 when it came to understaing refactoring... I found this [blog post](https://www.bmc.com/blogs/code-refactoring-explained/#:~:text=Code%20refactoring%20is%20defined%20as,sometimes%20known%20as%20micro%2Drefactorings.) to be very helpful as it explained why we use refactoring. If I went off the timing code we ran I would just assume that it is to make the macro run faster, much faster actually. But it is there to do more. 
### Advantages of Refactoring
* Cleaning up _"dirty"_ code will help to make the code easier to read for other coders, when things are clear and consise it is always easier to read and take away information. 
* It can then serve as a base for other coders. 
* When code is refactored it is easier to update and modify later. 
### Disadvantages of Refactoring
*  In some cases it may be more helpful to not refactor, sometimes leaving some additional steps may help you as the code writer understand each step
*  If there are a few different tasks that one macro is performing it can be helpful to have it broken down, that may help you to build the code in different ways or add to what the macro is performing. 

### If just looking at excel for the advantages and disadvantages of refactoring, I would say they are basically the same as with general coding. The main difference with gerneral coding and VBA in Excel is that with VBA you can only create code to run with microsoft office where as with general coding it can be performed on multiple platforms. An advantage to refactoring would be that the lines of data are clearer, with excel the spacing does not impact the code so it may be more helpful to make sure you are "cleaning up" your code to avoid duplications that may slow down your code. A disadvantage would be that it can be difficult to format VBA code, VS Code makes it easier to visulize the way the code is running and without this, when you refactor you may actually alter the way the macro is running which may cause you to spend more time trying to debug code that was once functioning. 


