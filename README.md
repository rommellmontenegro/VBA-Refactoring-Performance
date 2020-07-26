# Refactor VBA Code and Measure Performance

## Overview of Project
    *The dataset included stock ticker symbols, trading volumes, and prices of those stocks at open,
    high, low, close, and adjusted close for 2017 and 2018.

## Purpose
    *This challenge employed a wide range of VBA coding to speed up the automated calculation of the
    aggregating volumes and 12-month change in price per stock ticker. By properly refactoring the code,
    the time it takes to run the script will shorten.

## Analysis and Challenge
    * The main challenge inherent to this exercise is to introduce more complexity in the structure of
    the code that is written. This was done by adding arguments which are more abstract. For loops,
    conditional statements, arrays, formatting, and variable types within VBA's object-oriented coding
    permits the development of a useful macro.
    *This macro performs many calculations by manipulating the data in a table to produce another data
    table with useful calculations about the performance of the stocks.

## Results
    *2017 was a much better year in terms of the increase of stock price values than 2018.
    *For all but 5 stocks, for each stock the volume sold in 2018 exceeded the volume sold in 2017.

## Advantages and Disadvantages of Refactoring Code, Generally
    *There are many ways of creating a VBA script that can output a series of calculations and
    present them in a table. Some are more efficient than others.
    *The use of nested loops in programming is a powerful way of instructing a computer to iterate a
    series of commands perfectly and quickly. However, designing code which requires more system
    resources will eventually drag on a computer's performance if the amount of data analyzed is
    greatly increased or if the number of calculations required increase.
    *The design of efficient code will take advantage of producing all of the required calculations
    in an order which can read and use the data once to simultaneously report any calculations.
    *Arrays are instrumental to efficiently storing the values that a script's calculations produces.
    The use of arrays enables a computer to focus on the processing of data as it iterates through a
    data set. Once completely populated that same array is cleverly used to report the calculations.
    *A con of refactored code is that the script can get more abstract. More abstract code
    is more difficult for someone else who is new to it to read/understand quickly.
    *Another con of refactored code is that use of more abstract code requires more thoughtfully
    written comments to explain the data processing or the arguments employed in the script
    *If the calculations or the subject of the data being analyzed is already complex, then introducing
    advanced programming could make subsequent changes to the code more difficult to administer since
    the code becomes more byzantine.
    *More byzantine code is harder to explain quickly or effectively in writing or verbally.
    *Refactored code should make changes that only introduce the minimally required deviance or
    complexity from the original code.

## Advantages and disadvantages of this original and refactored VBA script
    *After refactoring this code, the time it took to produce either set of results for 2017 or 2018
    was diminished by 9X.
    *The 2017 data took 2.71875 long seconds before refactoring while it only took
    a fraction of a second, .3125 seconds, afterward.
    *The 2018 data took 2.820312 long seconds before refactoring while it only took
    a fraction of a second, .3203125 seconds, afterward.
    *The main difference before and after the code's refactoring was that the script only read the data
    once to produce the calculations after refactoring its design.
    *The additional arrays in the refactored version of this code required more variables to represent
    indices of the rows within the data self and the indices of the new arrays. Differentiated array
    nomenclature became necessary in order to more clearly represent which index was being used.

#Images
    *Please see images for this challenge within the GitHub folder:
