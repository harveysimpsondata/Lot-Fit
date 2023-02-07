# Lot Fit Matrix

## Problem Statement
- A company I used to work at was facing the challenge of performing lot fits manually.  The client would provide a lot fit matrix containing the entire house lineup, and the company had to calculate the backyard length for each footprint and only display the footprints that would be sold within the project.

## Solution
- To solve the problem, I proposed a solution to find the largest rectangle that could fit within the building setbacks.  I utilized a LISP solution from the AutoCAD forum (https://forum.autodesk.com/t5/visual-lisp-autolisp-and-General/draw-largest-rectangle-within-a-limited-shape/td-p/9399866) to automatically draw the largest rectangle.  I then drew a polyline from the middle of this rectangle to the back of the lot and extracted the relevant data into a CSV file.  This data was then pasted into the Lot Fit Matrix.xlsx on the 'Length' sheet.

- Once the data was set in the 'Length' sheet, I ran the python file 'LOT FIT STEP 1.py', which calculated the backyard length in feet for each building design, and highlighted cells in red if the house wouldn't fit within the setback.  This script produced a new Excel file called '1_Lot Fit Matrix with Exposed Columns.xlsx'.

- The python file 'LOT FIT STEP 2.py' was then run, and the client's desired building designs were entered into the list within the script.  This list would hide all columns that were not within the list, and the option for "Side = True/False" would display building designs with side garages if True, or hide the columns if False.

- After which the python file 'LOT FIT STEP 2.py' would run.  The client would give us certain building designs that they wanted to use for each subdivision so I could fill in the list within 'LOT FIT STEP 2.py'.  This list hides every column that is not within the list.  There is also an option for Side = True/False.  This option would show the building designs with side garages if True and hide the columns if False.

## Savings
- The manual process of performing lot fits previously took 3 employees 3-4 days per job. With my system in place, it reduced the time to 1.5 days per job and saved the company approximately $8000-$10000 per job.
