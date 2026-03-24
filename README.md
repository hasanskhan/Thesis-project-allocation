# Thesis-project-allocation
A MILP model that allocates projects based on preferences of participants. 

**Libraries used:**
scipy — the core of the allocation engine. Specifically scipy.optimize.linear_sum_assignment which implements the Hungarian algorithm — this is the mathematical solver that finds the globally optimal student-project assignment. This is the equivalent of what glpsol (GLPK) was doing in your original AMPL code.

numpy — used to build the cost matrix (the grid of students × projects with preference scores). scipy depends on it anyway so it's always available alongside scipy.

openpyxl — handles all the Excel reading and writing. Reads your survey .xlsx file and writes the formatted results file with colours, borders, frozen rows etc.

re — Python's built-in regex library, used to extract project codes (e.g. C037) from cell text. No installation needed.

datetime, sys, collections — all Python built-ins, no installation needed.
