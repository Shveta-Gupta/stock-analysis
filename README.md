# stock-analysis
Learning VBA Macros by analyzing stock data
## Overview of Project
#### Using VBA and a dataset containing stock information from the years 2017 and 2018, we were able to create a script that analyzes the yearly returns and daily volume of the stocks to determine what would be the best investment. To create this script, nested for loops were utilized to easily loop throuh all the data points and call to different aspects of the dataset without writing too many lines of code.

## Results
### Stock Performance
#### The stock performances for 2018 and 2017, respectively, are shown below: 
![image](https://user-images.githubusercontent.com/58957855/124296897-7733f280-db28-11eb-8319-d418945082f0.png)
![image](https://user-images.githubusercontent.com/58957855/124296939-861aa500-db28-11eb-8352-95331377bbfc.png)
#### As clearly shown by the color of the yearly return cells, the stock perfromance for 2017 was much better than 2018. The yearly returns are coded so that anything below 0% would be red, and anything above would be green -- indicating that green shows and increase in value and a better investment choice. 
### Timing of refactored code.
#### Below are the results of running analysis for all 2018 stocks. The first image is the original script, and the second is the refactored.
![image](https://user-images.githubusercontent.com/58957855/124292260-616ffe80-db23-11eb-82ec-44b0a17159e6.png)
![image](https://user-images.githubusercontent.com/58957855/124292463-97ad7e00-db23-11eb-89d6-58a6c3eac5ae.png)
#### As you can see, the refactored script ran much faster, saving almost an entire second. While this is a small differnce, if this difference was amplified by using a much larger dataset, the refactored code timing would be much more important.

## Summary
### Advantages and Disadvantages to refactoring code
#### Refactoring code allows programs to run faster, and makes writing code more adaptable. However, if the code you are refactoring from has a major design flaw, reusing that code can cause an even larger mess. 
### How do the pros and cons apply to refactoring the original VBA script?
