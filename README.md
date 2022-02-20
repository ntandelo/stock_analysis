# Stocks Analysis with VBA

## Overview of Project
Client wants to expand his research to include the entire stock market over multiple years. Code has been refactored to loop through all data one at a time to collect information per year, as input by user. 

### Purpose
To analyze stock performance for selected stocks over years; to evaluate potential future performance based on past performance. 

## Results Summary

![image](https://user-images.githubusercontent.com/98437495/154866309-48b3c3b4-4e4a-4058-a8e7-44ae681c8a85.png)
![image](https://user-images.githubusercontent.com/98437495/154866318-8e297e91-4924-4de6-af98-d509f8221885.png)
### Analysis of 2017
In 2017 only TERP showed a negative earnings over the year, all other stocks grew. 

### Analysis of 20178
In 2018 most stocks did not perform well; only ENPH and RUN showed positive gains.

## Conclusion
While stocks generally trended down in 2018, it is difficult to say whether this is related to individual stock performance or the stock market as a whole. Given that most stocks showed a decrease, it can be inferred that the market as a whole in 2018 was not good for stock values across the board. 

## Advantages and disadvantages of refactoring

Code refactoring is very useful if the programmer already has a skeleton or other functioning code that requires simple augmentation and additions in order to produce a wider array or a more complete data set of results. It is very useful to save time and effort when running code, and makes a lot of sense in the real world since often clients will return with projects that are just additions or extensions to a product you have given them already.

The only disadvantage would lie in user error. If the code is not well commented or is disorganized and cannot be easily followed, it will be difficult and time consuming to refactor effectively. 

### Refactoring of this VBA Script, speficially
It was important to refactor this script since in the first iteration of this code only 2018 was being evaluated, so half our data was simply ignored. In this refactor, the code ran faster than before. 
