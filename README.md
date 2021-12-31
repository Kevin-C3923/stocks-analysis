# stocks-analysis

## Overview of Project
  Establishing a VBA script that would be usable when running hundreds/thousands of stocks and user friendly for non coder to be able to use the VBA script.

## Results

  When performing the outcome we could see a difference when running the code for both 2017 and 2018. When using the pop up of the execution times, 2017 run time with the orginal was about 0.70~ seconds while the refactored time was about 0.14~ secomds

![2017-Pic](https://github.com/Kevin-C3923/stocks-analysis/blob/main/Resources/VBA_Challenge_2017.jpg)

  This can be said for 2018, as orginal was 0.80~ seconds to 0.15~ seconds to refactored.

![2018-Pic](https://github.com/Kevin-C3923/stocks-analysis/blob/main/Resources/VBA_Challenge_2018.jpg)

  The difference between these times has to do with the code we changed, which was the two For statement or a double For Loop. The double For Loop is a simple code to use when producing/using small output, but when you start increasing those output the time to produce results will take longer to finish. Therefore, removing the For loop that is inside the other will greatly increase the time to complete the assignment.  

## Summary
### Advantages/Disadvantages of Refactoring Code
  The purpose of refactoring code is to improve on a code while still doing the same thing. In this case, we slightly different from each other, so you would assume their was no purpose in facatoring. However, we are just using a small portion of stocks compared to thousands that are out their. So when we have to put this code up against more stocks then our efforts will be rewarded. 

  But, one of the major obstacle is what needs to be removed and what needs to stay in the original code. We have to make sure that it still does what it orginal did, find the Volume and its Return. Therefore, we have to carefully follow through the code to determine this, which is time consuming. Additionally, even if we remove a piece of the code we have to make sure that the new line of code is a better replacement and those what we want it to do.

### Pros/Cons of refactoring original VBA script
When working on the refactoring of the original VBA script, many parts of the code were not changed like the code for the timer, editing the outlook of the table, deleting the table and the original ticker code. The code that had to be change was the  double For loop code, as this was the part that would make the script run longer when it in take more stocks. In order to run faster we would have to remove the For that is inside the other For loop, but we need to have other arguments to produce the similiar output. In this case we will have to use multiple arrays and variable to store the information. Additionally, the If Then statement also need changes by using the new variable as it can't use the old For loop info. This will take time as many errors will arise due to the small changes that has occured, but once your able to determine which variables and arrays goes where the code should produce the same outcome as the original and run a bit faster.
