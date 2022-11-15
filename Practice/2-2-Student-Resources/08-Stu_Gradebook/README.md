# Gradebook

## Instructions

* Using `grader.xlsm` as a starting point, create a grade calculator using **conditionals**. This calculator will convert a student's numeric grade into a letter grade, and style the resulting cell accordingly.

* Once complete your script should perform the following:

  * If the score is over 90, the student will receive an "A" in the letter grade cell, and the Pass/Warning/Fail cell will be filled green with the text "Pass."

  * If the score is between 80 and 89 (inclusive), the student will receive a "B" in the letter grade cell, and the Pass/Warning/Fail cell will be filled green with the text "Pass."

  * If the score is between 70 and 79 (inclusive), the student will receive a "C" in the letter grade cell, and the Pass/Warning/Fail cell will be filled yellow with the text "Warning."

  * Finally, if the score is below a 70, the student will receive an "F" in the letter grade cell, and the Pass/Warning/Fail cell will be filled red with the text "Fail."

## BONUS

* Create a second button that resets the grades to the original state and then establishes the previous grade in a row labeled "Last Grade."

* Hint: the built-in color value for no fill is 'xlNone'

## Answer
Sub Grader()

'Check the students grade is greater than or equal to 90
If Cells(2, 2).Value >= 90 Then

    'Est. the grade is passing
    Cells(2, 3).Value = "Pass"
    
    'Color the passing grade green
    Cells(2, 3).Interior.Color = vbGreen
    
    'Set the letter grade to A
    Cells(2, 4).Value = "A"
    
'Check if the students grade is great than or equal to 80
ElseIf Cells(2, 2).Value >= 80 Then

    'Est. the grade is passing
    Cells(2, 3).Value = "Pass"
    
    'Color the passing grade green
    Cells(2, 3).Interior.Color = vbGreen
    
    'Set the letter grade to B
    Cells(2, 4).Value = "B"
    
'Check if the students grade is greater than or equal to 70
ElseIf Cells(2, 2).Value >= 70 Then

    'Est. the grade is "Warning"
    Cells(2, 3).Value = "Warning"
    
    'Color the warning grade yellow
    Cells(2, 3).Interior.Color = vbYellow
    
    'Set the letter grade to C
    Cells(2, 4).Value = "C"
    
'Check if the students grade is less than 70
ElseIf Cells(2, 2).Value < 70 Then

    'Est. the grade is "Fail"
    Cells(2, 3).Value = "Fail"
    
    'Color the warning grade red
    Cells(2, 3).Interior.Color = vbRed
    
    'Set the letter grade to "F"
    Cells(2, 4).Value = "F"
    
End If
    
End Sub


Sub Reset_Grader()

    'Pass the previous grade to the Last Grade row
    Cells(12, 2) = Cells(2, 2).Value
    Cells(12, 3) = Cells(2, 3).Value
    Cells(12, 4) = Cells(2, 4).Value
    
    'Empty out the current grade and remember to set the color back to default
    Cells(2, 2).Value = ""
    Cells(2, 3).Value = ""
    Cells(2, 3).Interior.Color = xlNone
    Cells(2, 4).Value = ""


End Sub

