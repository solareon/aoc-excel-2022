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