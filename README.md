<div align="center">

## paramarray and function replacement


</div>

### Description

1. Shows you how to use paramarray to create much more flexible and capable functions

2. Shows you how to REPLACE vb's existing functions
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Izek](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/izek.md)
**Level**          |Intermediate
**User Rating**    |4.8 (120 globes from 25 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/izek-paramarray-and-function-replacement__1-28718/archive/master.zip)





### Source Code

<br><br>
In this tutorial I will show you how to write your a function that can accept ANY number of parameters and how to replace existing vb functions so that when you call function Right for example it will execute YOUR version of the function instead of vb's version.
<br><br>
As some of you may or may not know VB has a function called SWITCH
<br>
Function Switch(ParamArray VarExpr() As Variant) As Variant
<br>
That function evaluates a list of expressions and returns a Variant value or an expression associated with the first expression in the list that is True. Meaning you can do something like. . .
<br><br>
Dim i as integer
<br>
Dim retval as boolean
<br>
i = 1
<br>
retval = Switch(i = 1, True, i = 2, False)
<br><br>
When you execute that it will return true because i is 1, if i was 2 it would return false, but if i was 3 it will ERROR!!!!!!!!! because none of the expressions evaluated to true.
<br>
Also another thing about this function is that you can pass as many parameters as you want and that is what makes it so special for our purposes.
<br><br>
What I decided to do was to REPLACE VB's existing switch function with my own switch function so that when i call the switch function and none of the expressions evaluate to true it will either return "" OR it will return the "default parameter".
<br><br>
Now, to explain how it works and what the default parameter is . . .
<br><br><br>
Function Switch(ParamArray VarExpr() As Variant) As Variant
<br>
'paramarray makes it so that you can pass as many parameters as you want which can be accessed using the VarExpr array.
<br>
Dim i As Integer
<br>
For i = 0 To UBound(VarExpr) Step 2
<br>
'this loop will go through every other argument in our parameter array also note, when we pass argument like i = 1 VarExpr for that argument will not be "i = 1" it will be True if i is 1 or it will be False if its not
<br>
If VarExpr(i) = True Then
<br>
'check to see if argument evaluated to true
<br>
Switch = VarExpr(i + 1)
<br>
'return the value for that argument
<br>
Exit Function
<br>
End If
<br>
Next i
<br>
'if none of the arguments evaluted to true this part will check if you have even or odd number of parameters if you have ODD number of parameters it will assume that the last parameter is the default parameter which is to be returned if nothing evaluted to true
<br>
If (UBound(VarExpr) + 1) Mod 2 = 1 Then
<br>
Switch = VarExpr(UBound(VarExpr))
<br>
'return the last ("default") paramter
<br>
End If
<br>
End Function
<br><br>
Also note that our function name is Switch just like VB's function name, we put our function in a module so that when you call Switch function it will call your(more flexible) version of the function and not VB's default version.
Now when we call our new function . . .<br><br>
Dim i as integer
<br>
Dim retval as string
<br>
i = 1
<br>
retval = Switch(i = 1, True, i = 2, False)
<br><br>
retval will be "True", but if we take out i = 1 it will return ""
<br><br>
Dim i as integer
<br>
Dim retval as string
<br>
retval = Switch(i = 1, True, i = 2, False)
<br><br>
because i is 0 and none of the expressions evaluted to true, but what we can do is add the extra "default" parameter<br><br>
Dim i as integer
<br><br>
Dim retval as string
<br>
retval = Switch(i = 1, True, i = 2, False, Now)
<br><br>
now when the function will execute it will return current time and date (because thats what function now returns) because none of the expressions evaluated to true
<br><br><br>
If you do not understand the above explation and want a more indetail explanation and/or example you can contact me at
<br><br>
email: izek.programmer@verzion.net
<br>
aim: ozik13
<br>
icq: 53982424
<br>
Please leave feedback :)

