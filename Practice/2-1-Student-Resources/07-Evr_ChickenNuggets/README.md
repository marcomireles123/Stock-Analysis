# Looping on through

Now it's your chance to see how quickly we can create data utilizing the power of a computer and `for` loops!

## Instructions

* Create a `for` loop that will produce the following example. The lines signify new cells.

|  A | B  |  C |
|:---:|:---:|:---:|
| I will eat | 11 | Chicken Nuggets |
| I will eat | 12 | Chicken Nuggets |
| I will eat | 13 | Chicken Nuggets |
| I will eat | 14 | Chicken Nuggets |
| I will eat | 15 | Chicken Nuggets |
| I will eat | 16 | Chicken Nuggets |
| I will eat | 17 | Chicken Nuggets |
| I will eat | 18 | Chicken Nuggets |
| I will eat | 19 | Chicken Nuggets |
| I will eat | 20 | Chicken Nuggets |

## Answer code:
* Sub loops_and_loops()

    * Loop through the first 10 rows
        * For i = 1 To 10
    
        * 'Set values in coulmn A to "I will eat"
        * Cells(i, 1).Value = "I will eat"
        
        * 'Set values in column B to the sum of the counter + 10
        * Cells(i, 2).Value = i + 10
        
        * 'Set values in column C to "chicken nuggets"
        * Cells(i, 3).Value = "chicken nuggets"
        
    * 'Call the next iteration
    * Next i

* End Sub

