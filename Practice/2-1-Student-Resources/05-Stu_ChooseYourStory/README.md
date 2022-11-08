# Choose Your Story

## Instructions

* Create a simple Excel workbook and VBA macro in which a user is provided a single button to click. Based on the number they provide in a text box above, a different message box will appear.

  * If the user enters a value of 1, display: “You choose to enter the wooded forest of doom!”
      * If (Range("B1").Value = 1) Then
        MsgBox ("You choose to enter the wooded forest of Doom!")

  * If the user enters a value of 2, display: “You choose to enter the fiery volcano of doom!”
      * ElseIf (Range("B1").Value = 2) Then
        MsgBox ("You choose to enter the fiery volcano of Doom!")

  * If the user enters a value of 3, display: “You choose to enter the terrifying jungle of doom!”
      * ElseIf (Range("B1").Value = 3) Then
        MsgBox ("You choose to enter the terrifying jungle of Doom!")

  * If the user enters a value of 4, display a similar custom message.
      * ElseIf (Range("B1").Value = 4) Then
        MsgBox ("You stay at home and code for fun!")
  * If the user enters anything else, display: “Try following directions”
      * Else
        MsgBox ("Try following directions!")