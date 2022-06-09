# stock-analysis

#Stocks Challenge

Just as the title of this project states, the goal was to conduct stock analysis. For this project we focused on data from 2017 and 2018. The stocks are listed by their ticker symbols, which are in the A column of sheets "2017" and "2018. Using output arroys, for loops and if else statements, we created a VBA code that returns the results of the stocks. In addition to returning the results per year, it also color codes them in red and green to reflect negative(red) and positive(green) return values per ticker. 

**Analysi**

The first thing we do is create a ticker index and set it equal to 0 since that is the value that VBA uses as the starting index. We then create three arrays; tickerVolumes, tickerPrices, and tickerEndingPrices and use types long and single. The next step is to initialize the tickerVolumes and we make this equal to 0. We start at 0, so that we can increase the tickerVolumes based on the volume and the previews tickerVolume. Next step was to make sure that we know where the first and the last rows of data are in the tables that we are working with. Finally, since the end of our if statement is coming up, we increase the tickeIndex by 1 and let it loop for the next values. 

The fourth step of this was to print out the "All Stocks Analysis" sheet within the worksheet. This was done by printing the tickers(i) for the first column, tickerVolumes(i) for the second and finally tickerEndingPrices(i) / tickerStartingPrices(i) - 1 for the last. This is another for loop, so it would look through the arroys and stop ove it has reached the end of the referenced table. 

**Results**

When we run the code for years 2017 and 2018, we get the popup window that we created earlier. For year 2017 we get the message "This code ran in 0.1328125 seconds for the year 2017". For the 2018 year, "This code ran in 0.140625 seconds for the year 2018. We can see from this that the 2017 code ran faster than that of the 2018. We also recognize that 2017 had all but one positive return values, while 2018 has all but 2 negative return values. This result indicates that 2017 investors earned more in returns than those who invested in 2018. The code for 2017 also ran faster than the one for 2018 and that could be because of the negative values and the for loops taking longer for those. ![VBA_Challenge_2017.png"](https://user-images.githubusercontent.com/92186586/172767732-4df0a2f3-a178-461d-9492-7a422bc98244.png">)

**Summary**

The advantage of refactoring in general is that it makes the code easier to understand. It is more organized and clean looking. One of the disadvantages is that the original writer of the code may not be the same person as the one reusing this. In this case, it could be confusing to try to understand the meaning behind variables and the purpose of the code. There could be some code bugs that were solved after the original code that may now reappear. 

Refactoring in VBA can also be tricky which is a disadvantage. Although an advantage is that the code looks cleaner and easier to understand, there could be bugs from misplacing certain code. For instance, if the for loop includes if/else statements with misplacement of the elif, there could be a possible infinite loop. 
