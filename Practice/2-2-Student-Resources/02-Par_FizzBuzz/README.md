# Fizz Buzz

## Instructions

* Create a VBA Script that populates the second column with the word "Fizz", "Buzz", or "Fizzbuzz" based on the value in the first column.

  * If the value in column 1 is a multiple of both 3 and 5, print "Fizzbuzz" in column 2.

  * If the value in column 1 is a multiple of just 3, print "Fizz" in column 2.

  * If the value in column 1 is a multiple of just 5, print "Buzz" in column 2.

## Hints

* Remember the mod!

## Answer
Sub Fizzbuzz()
    'Loop through the values in Cln. 1
    'Rows 2-100 are the rows we want to manipulate
    For i = 2 To 100
    
        'Set cell value variable
        num = Cells(i, 1).Value
        
        'Check if the # is divisible by 3 + 5
        If (num Mod 3 = 0 And num Mod 5 = 0) Then
        
            'If so print Fizzbuzz
            Cells(i, 2).Value = "Fizzbuzz"
            
        'Check if the # is divisible by 3
        ElseIf (num Mod 3 = 0) Then
        
            'If so print Fizz
            Cells(i, 2).Value = "Fizz"
            
        'Check if the # is divisible by 5
        ElseIf (num Mod 5 = 0) Then
        
            'If so print Buzz
            Cells(i, 2).Value = "Buzz"
            
        End If
    
    Next i
        
End Sub

