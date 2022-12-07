# aoc-excel-2022
Excel is a still programming language. This repo archives my attempt to do all of the Advent of Code 2022 entirely in Excel without the use of VBA.

# Day 1
Sum until you hit a blank then write a 0 and repeat. Part one looks at MAX() over the range, Part 2 looks for the top 3 by using a shortcut and assuming that no single elf will have a maximum between two values of another elf. 

# Day 2
Remove the space to make typing the SWITCH() cases out easier and then build the truth tables for part 1 and part 2. Observant folks will see that each combination can only correspond to a single number between 1 to 9 with no repeats.

# Day 3

Part 1 is solved in a single cell with this monstrosity which splits the string evenly and then into a by character array and use FIND() as a case sensitive search across the array. This takes advantage of the fact that inputs on a single row don't repeat characters across the middle. After finding the character morph it into the ascii code with CODE() then shift the value down depending on if it was above 96 or below 96

Part 2 almost threw this idea out the window but worked in a similar fashion. First operation is to split the string into a by-character array then convert to ascii code as UNIQUE() is not case sensitive in Excel. Then combine the three arrays with HSTACK() and then converted to ascii and sorted again then CONCAT() back to a single string. An array of all possible 3 character upper and lowercase combinations is built with REPT() and fed into FIND() to find the starting position of that which is returned with MID(), converted to ascii and then shifted same as part 1. Figured out how to get it down to a single cell to solve part 2

# Day 4

Laziness wins, split the text into 4 columns, do some IF() AND() OR() magic and then sum the resultant columns. TEXTSPLIT() and NUMBERVALUE() are important to remember as excel will try to compare numbers stored as text.

# Day 5

Switch statements go BRRRRRR. Formatted input across some helper cells the pulled columns into single cells. Another helper function parses the instructions down to single numbers. A giant LET() formula then glues the whole mess together as you fill down to row 513. Grab the right character from each cell of the final array and smash together.

Part 2 does the same thing but removes the ReverseText() lambda from the move part.

The demo sheet has the table that assembles the SWITCH() cases programmatically since I kept getting tripped up making it. Some concat() and textjoin() magic there glues that all together.

# Day 6

Surprisingly short solution to this one today. Both Part 1 and Part 2 used the same formula with two changes to the offset.

First step was to walk the input and break it into 4 or 14 character chunks, then convert to ascii code values in an array and unique() the array, convert back to characters and see the length, if it’s 4 or 14 then it is a match and we search for the string in the input and add 3 or 13 since we found the starting position of the message. I could probably shorten the solution and combine to a single cell for both parts with a let but that’s a later task when I’m at a real computer not on the iPad.

# Day 7

What an awful thing to reconstruct in Excel. First step is to use TEXTSPLIT() to pull out the data into two columns, one with the command or file size and the second with filename or directory name. Next step is to use some hideous formula to reconstruct the full paths respecting the .. to go up the tree. Next column is a simple extraction of the numbers and formatting to allow Excel to compute them. Next step is to pull the unique directory tree, sum at that level and then build a lookup table to find subdirectories under that directory. The lookup table is summed across a second row to find the real total size of a directory. Part 1 is solved by SUMIF() across this data set. Part 2 requires a little bit of trickery to find the target number. The input list is the same as part 1 but uses a combination of ABS() and MIN() to find a value that fits the critera.