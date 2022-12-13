# VBA-challenge

## Process

Note: I left all of my previous attempts at the code that are hashed out so that the full process is visible.

First I defined variables that I would be using as the type of value they are. 

I also defined column and row titles early so that they'd be ready to be filled in later on. 

Then I set the variables for the different rows that I'd be using. I needed one for the general last_row of a column, one for the summary_row, where the output would be printed, and one for finding the first row of a ticker to determine the open_row.

I set the variables I needed to 0. 

The first for loop was for multiple parts of the challenge. Firstly, it found the ticker value and calculated the total change in value by finding each ticker's closing value and subtracting its opening value from it. That value was placed in column K, under the variable 'change'. 

Percent was similarly calculated by dividing 'change' by 'open_' and those values were stored in row L, formatted to a percentage. 

Total ticker volume was stored in column M with variable 'volume_total'.

The ticker name was stored in column J. This was found by checking to see if the value of a cell was different from the one above it. 

Else, the volume of a row was added to 'volume_total'.


The next for loop changed the color of positive yearly changes to green and the ones that were negative to red. 

The following for loop set the 'greatest' variable in order to find the greatest % change. That value from column L would then be placed into the table at R2, formatted at a percent. The ticker name would be printed in the adjacent cell Q2. A similar for loop was added to find the 'lowest' % change and the values found were placed in Q3 and R3. 

The last for loop was very similar to the one that found the greatest % change, but was used on column M to find the greatest total stock volume 'greatest_vol'. 