# Stock Analysis

## Overview
The purpose of this analysis was to refactor the VBA code in order to optimize the performance. The previous iteration of the worksheet was designed using only two datasets and there were concerns about its ability to handle larger amounts of data. The refactored code should perform faster allowing for larger datasets.

## Results
### Stock Performance
When comparing the two years, 2017 was clearly the top performer with all but one ticker bringing in a positive return. The one ticker that was in the red managed to keep it below a ten percent loss.

![2017-stocks-module-results](https://user-images.githubusercontent.com/89947873/147859741-87724eba-3579-4b09-9b50-688ea458dd98.png)

The performance in 2018 was red across the board with two exceptions who both posted impressive gains.

![2018-stocks-module-results](https://user-images.githubusercontent.com/89947873/147859776-7baa6836-6c35-45ca-b82d-f372bc327119.png)


### Code Performance
The performance of the code saw a definite improvement. These are the original run-times for:

2017

![DQ-stocks-2017](https://user-images.githubusercontent.com/89947873/147859801-9b5fd348-b54c-4221-823d-bc19c3362dc2.png)

and 2018

![DQ-stocks-2018](https://user-images.githubusercontent.com/89947873/147859804-3f564cc9-e01c-41cd-aaa0-012f452ddbc6.png)

Compared to the refactored run-times:

2017

![VBA_Challenge_2017](https://user-images.githubusercontent.com/89947873/147859812-91cc26e0-5945-40f4-b201-63aa7def8dff.png)

and 2018

![VBA_Challenge_2018](https://user-images.githubusercontent.com/89947873/147859813-db6cea55-4dcb-40c3-ad76-6d9751eec28d.png)


## Summary
There are many advantages, as well as disadvantages, when it comes to refactoring code. One of the disadvantages is in the diversity of coding possibilities. There are often many different ways to solve the same problem. These variations can not be compared without exhaustive testing and comparison. Don't forget the advantages. Refactoring code is an easy way to make it cleaner and easier to read which improves the overall design. Also, as we saw, refactoring increases the processing speed as well. This increases the programs ability to handle larger amounts of data as well as its overall usefullness. 
