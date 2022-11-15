# Star Counter

## Instructions

* Create a VBA Script that tallies the number of "Full Stars" per row and enters them into the Total column. Starter Code is provided, but feel free to start from scratch if you want an extra challenge :-)

* **Hint:**

  * You will need to use a nested for loop.

  * You will need to create a variable to hold the number of stars and continually reset this variable at the start of each row.

* **Bonus:**

  * **Part 1:** Automatically determine the last row. 
  
    * Instead of hard-coding the last number of the loop, use VBA to determine the last row automatically (i.e. do not use for i = 2 to 51)

  * **Part 2:** Visualize the Results 

    * Using a Pivot Table, determine if there is a relationship between Review Date and Rating using a line chart.
    
    * Using a Pivot Table, determine if there is a relationship between Program Type and Rating using a bar chart.

* Good luck!

## Answer
Sub StarCounter()

  ' Create a variable to hold the StarCounter. We will repeatedly use this.
  Dim StarCounter As Integer
  
  ' BONUS: counts the number of rows
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
 
  ' Loop through each row
  ' BONUS: Use lastrow variable instead of 51
  For i = 2 To lastrow

    ' Initially set the StarCounter to be 0 for each row
    StarCounter = 0

    ' While in each row, loop through each star column
    For j = 4 To 8

      ' If a column contains the word "Full-Star"...
      If (Cells(i, j).Value = "Full-Star") Then

        ' Add 1 to the StarCounter
        StarCounter = StarCounter + 1

      End If

    Next j

    ' Once we've completed all rows, print the value in the total column
    Cells(i, 9).Value = StarCounter

  Next i
  

End Sub

