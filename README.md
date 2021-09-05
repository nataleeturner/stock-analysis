# stock-analysis

## Overview of Project

The purpose of this project is to recommend a stock to invest in based on the total trading volume and the return of 12 different stocks during a specific year. In order to perform this analysis, VBA in excel was utilized to write a code that will output this information for the 12 stocks. Once the initial code was written, it was refactored to make it faster and more concise.

## Results

In 2017, DQ, ENPH, FSLR, and SEDG all had an excellent yearly return (>100%). FSLR had the highest traded volume at ~680MM, ENPH and SEDG each had about a third of that, and DQ was much lower at ~35MM. In 2018 only ENPH and RUN had a positive return year. Therefore, based solely on this information, since ENPH did well comparatively in both 2017 and 2018, I would recommend investing in ENPH. To be more confident in this recommendation, I would propose performing a more thorough financial analysis of the companies.

<img width="239" alt="2017_Stocks_Analysis" src="https://user-images.githubusercontent.com/88349443/132112789-0df701b9-fb5d-48f5-a130-73e71dd77888.png"> <img width="233" alt="2018_Stocks_Analysis" src="https://user-images.githubusercontent.com/88349443/132112793-d2a9bc50-9265-4573-8e0f-5682c5296753.png">

The original code ran in about 0.30 seconds for each year and the refactored code runs in about 0.07 seconds, or 77% faster. See below the reduction in run time between the run times of the two scripts. 

<img width="263" alt="Original_Script_2017" src="https://user-images.githubusercontent.com/88349443/132112811-c528b84a-7346-4fe9-b44e-0bc1007d68b5.png"> <img width="262" alt="Original_Script_2018" src="https://user-images.githubusercontent.com/88349443/132112817-3401d721-9939-490f-8975-036c18729981.png"> 

<img width="265" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/88349443/132112822-068dfa9b-bda9-43f6-b176-af5289357d68.png"> <img width="267" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/88349443/132112824-0d2dec54-1f5a-4e21-bb64-ffefb12291e1.png">

### Removing nested for loops
One characteristic of the refactored code that makes it so much faster is that nested for loops were eliminated by adding in a new index called tickerIndex. This index allows for the removal of the first "for" loop, by making it so the index increases by one each time the second (in the original script) for loop is completed.

### Storing data in arrays
Another change that was made is to store all of the data into arrays and write it to the analysis tab one time at the very end. This makes it so that the code isn’t constantly switching back and forth between the data tab and the analysis tab with each loop. The first section of code below shows the initialization of the tickerIndex and the arrays used to store data that did not exist in the original script. The next two sections of code show the difference in the final ouput between the original (which is part of a nested loop) and the new (unnested loop) code.

<img width="253" alt="VBA_Challenge_Code_Snipit" src="https://user-images.githubusercontent.com/88349443/132112939-14cb5cdb-c356-4e98-9b95-584ab90b6f06.png">

<img width="531" alt="Original_Script_Output" src="https://user-images.githubusercontent.com/88349443/132112999-ba431205-5e2d-4bc6-9539-fa18a69ddd1e.png"> <img width="369" alt="VBA_Challenge_Script_Output" src="https://user-images.githubusercontent.com/88349443/132113001-248996f7-9b20-46ee-9aaa-6388defc5675.png">

## Summary

### Advantages and disadvantages of refactoring code
One advantage to refactoring code is that it can make your code run much faster. The more you do with it, the more advantageous this can be. Additionally, it can make it easier to apply the code to other similar situations by making it more adaptable.

The major disadvantage to refactoring code is that the benefit may not be worth the cost (time) associated with doing the refactoring. A quick cost-benefit analysis would be something worth doing prior to any code refactoring.

### Advantages and disadvantages of the original and refactored VBA script
Some advantages of the original VBA script are that it was quick to write and easy to read. However, it was slower and more difficult to increase the scale of the data it’s being used to analyze.

The primary advantage to the refactored code is that it runs 77% faster. In this case, the difference between 0.3 and 0.07 seconds is probably inconsequential, but on a much larger scale, such as applying the code to the entire stock market, a 77% improvement could have major positive impacts. Some disadvantages to the refactored code are that there are more variables so there is more room for error when developing the script and it can make it slightly more confusing for someone else to come in and read/edit the code. 
