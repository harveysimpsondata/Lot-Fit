## Lot Fit Matrix

### Problem
- A company I used to work at had the problem of doing these lot fits by hand.  The client would give us a lot fit matrix with all of their house dimensions. We had to calculate the length of the backyard and fill into their matrix what the length of the backyard was for every design.

### Solution
- Since the lot has a setback for each one.  I somehow had to find the largest rectangle that would fit within these setbacks.  I proposed the problem in an AutoCAD forum here: https://forums.autodesk.com/t5/visual-lisp-autolisp-and-general/draw-largest-rectangle-within-a-limited-shape/td-p/9399866. 

- I used one of the LISP solutions that would automatically draw in the largest rectangle.  I then drew a polyline from the middle of this box to the back of the lot.  Since the box and line were on separate layers, I extracted the data I needed into a CSV and then pasted the data into the Lot Fit Matrix.xlsx on sheet 'Length'. 

- Once the data was set in sheet 'Length' I would run the python file 'LOT FIT STEP 1.py'.  It would calculate the backyard length in feet for each building design if the house would fit within the setback.  If the house would not fit then the cell would fill as red. This python file produces a new Excel file called '1_Lot Fit Matrix with Exposed Columns.xlsx'

- After which the python file 'LOT FIT STEP 2.py' would run.  The client would give us certain building designs that they wanted to use for each subdivision so I could fill in the list within 'LOT FIT STEP 2.py'. This list hides every column that is not within the list. There is also an option for Side = True/False. This option would show the building designs with side garages if True and hide the columns if False.

### Savings
- Usually this would take about 3 employees at least 3 days per job.
- With my system implemented it would take me at most 1.5 days per job.
- Ultimately it would save the company $5000-$7500 per job.
