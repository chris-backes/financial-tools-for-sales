# Modified Macros for Profit Analysis of Billed Labor

Modified Excel macros used in the context of sales of billed labor to see margin and total profit over time.

Please Note: These were written at a time when I was teaching myself to code. The code itself is clunkier than it needs to be in many places, and any breach of naming conventions was out of ignorance. These were written approximately two years prior to being first uploaded to GitHub.

Each of these started out as a recorded macro but were rewritten and are broken into three basic groups.
-Integrates a different analysis with ancillary tools (not discussed or presented here)
-Initiates the goal seek applicaiton within Excel, where the input variables are not provided in the pop-up window as usual, but rather refer to specified cells
-Switch the basis of profit analysis from either an hourly rate or salaried rate, and from either a (normalized) hourly rate or from a total cost (an hourly rate plus benefits/taxes).
(a normalized hourly rate is the hourly rate that would equal the yearly compensation a salaried employee would make if they were to switch to an hourly rate AND use the maximum allow time off provided under a company's time off policy)

![excel sheet with macro buttons and accompanying data](./assets/excel-display.JPG)

##Integration Subroutines
If multiple lines of analsys are, or can be, used for further analyses, these subroutines enable the user to switch between the basis of analysis. Box colors are changed to highlight which basis is being set.

The subroutine first targets the individual cells in the bottom row of the pictured image, then rewrites what the cells are equal to, depending on the subroutine initiated. Thus, for instance, BIntegration starts at cell c15, and sets it to the matching information in the second line of analysis (contained in cell rows 10 and 11), then moves to cell d15 and sets it equal to the matching information, etc.

Then, the subroutine highlights the relevant cellgs in column g (in Bintegration, the merged cells of g10 and g11) and removes the highlighting in the other g cells (if any).

##Goal Seek Subroutines
In normal circumstances, the goal seek can only be done by input text into the pop up window. The subroutines initiate a goal seek, where the range and goal value are predetermined cells, useful in contexts where one cell is consistently targetted for a goal. There are six subroutines which target six diffrent cells, and each have their own cells designated at the place where the value of the goal is stored.

The subroutine first selects the cell that contains the value looking to be modified. Then it initiates a goal seek. The goal value is determine by the value in the cell immediately above the button.

##Cost Basis Subroutines
If multiple lines of analysis are used to calculate profit, these subroutines can allow for easy comparison when the basis for those analyses are disparate.

### PY_Hourly and PY_Salary

Both of these rubroutines rewrite one or the other to be a funtion of the other. The highlighted cell is the declared variable, and the unhighlited on is written as a functionf the other (factoring in total workable hours, which is 2080-minus the total allowable paid time off--holiday, sick, general leave--one receives in a year).

### BR_TotalCost_MarkUp and BR_Normalized_MarkUp

Both of these subroutines rewrite various formulas contain in the three major rows of analysis to use the total cost or the normalized hourly rate as the basis of analysis.
