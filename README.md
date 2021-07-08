# stock-analysis
## Overview of Project
### Purpose
- The objective of this project was to refactor Microsoft Excel VBA code to collect and display information reguarding stock prices and conduct analysis on whether they would be reliable investments. This analysis was originally conducted in a similar manner however was refactored to increase efficiency and decrease the computing time of the code.
## Results
### Analysis
- Initially, I began to code each section of the project via following directions given through the VBA Module. Integrating two for loops, I was able to produce the desired results within sub-par time.  I then returned to this code and refactored it by deleting one of the for loops and using arrays to store the data in a way that would allow me to loop through it once and produce the same desired results in a much more efficient time. As seen below, I initiated three arrays: "tickerVolumes()", "tickerStartingPrices()", and "tickerEndingPrices()", that could hold twelve elements. Using a for loop, I iterated through these arrays, updating the information via conditionals placed within these loops.   "![StocksAnalysisCode](https://user-images.githubusercontent.com/35403433/125000554-a9f55380-e01e-11eb-9c4e-58518a2b004f.png)
## Summary
### What are the advantages or disadvantages of refactoring code?
- Refactoring code is a common technique used by many professionals that brings about many advantages. One of which is simply restructuring the code, making it easier for another person to read and understand what the code is doing. Another great advantage that refactoring creates is more effiecient code. This means that refactored code usually is able to to be processed and ran in a decreased time span. There are some disadvantages, however, as some programs are simply too large to refactor (we wouldn't have proper procedures or test cases to conduct refactoring). Thus, refactoring code that belongs to a larger program could pose the risk of a range of errors occurring, such as simple syntax errors or even larger logic errors.

### How do these pros and cons apply to refactoring the original VBA script?
As mentioned above, refactoring the original VBA script allowed for a large decrease in run time. The original script, when executed, took approximately one second to run, while the refactored script ran in little over a fourth of that time.![StocksVBARunTime2017](https://user-images.githubusercontent.com/35403433/125001613-0d808080-e021-11eb-9a05-a03c4ba91791.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/35403433/125001622-140ef800-e021-11eb-9b15-6dec6bbea02f.png)
