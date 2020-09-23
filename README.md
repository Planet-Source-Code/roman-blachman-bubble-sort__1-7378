<div align="center">

## Bubble Sort


</div>

### Description

This code should explain the bubble sort !

Bubble sort is a sort algorithm that sortes string in very fast way. There are many sort algorithms but this is the best !
 
### More Info
 
This code needs only one parameter:

The input array (will be used for output also)

If you have problems running it try this code.

(I assume that you created a list box that his name is: lstItems)

Place this code where you want (for example: command button)

Dim inArray() As String ' Input array

ReDim inArray(lstItems.ListCount - 1) ' Redim the array to the size of the list count items

For ic = 0 To lstItems.ListCount - 1

inArray(ic) = lstItems.List(ic) ' Put all the values from the list box to the array

Next

cBubbleSort inArray ' Sort the array

lstItems.Clear ' Clear the list

For ic = 0 To UBound(inArray)

lstItems.AddItem inArray(ic) ' Put the sorted items from the array

Next

This sub won't return anything. It will change the input array


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Roman Blachman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/roman-blachman.md)
**Level**          |Intermediate
**User Rating**    |4.5 (36 globes from 8 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/roman-blachman-bubble-sort__1-7378/archive/master.zip)

### API Declarations

no apis.


### Source Code

```
Public Sub cBubbleSort(inputArray As Variant)
	Dim lDown As Long, lUp As Long
	For lDown = UBound(inputArray) To LBound(inputArray) Step -1
		For lUp = LBound(inputArray) + 1 To lDown
			If inputArray(lUp - 1) > inputArray(lDown) Then SwapValues inputArray(lUp - 1), inputArray(lDown)
		Next lUp
	Next lDown
End Sub
Public Sub SwapValues(firstValue As Variant, secondValue As Variant)
	Dim tmpValue As Variant
	tmpValue = firstValue
	firstValue = secondValue
	secondValue = tmpValue
End Sub
This is the same code but with explainations:
Public Sub cBubbleSort(inputArray As Variant)
	Dim lDown As Long, lUp As Long ' Two variables that will be used in the fors
	For lDown = UBound(inputArray) To LBound(inputArray) Step -1 ' One variable will go from the upper bound of the array
		For lUp = LBound(inputArray) + 1 To lDown ' and the second one will go from the lowest bound to the top
			If inputArray(lUp - 1) > inputArray(lDown) Then SwapValues inputArray(lUp - 1), inputArray(lDown) ' This line check if the value from the up-to-down for is higher than the value from the down-to-up for, if so the sub call a swap sub that switches the values places
		Next lUp ' Continue to the next value from down-to-up
	Next lDown ' Continue to the next value from up-to-down
End Sub
Public Sub SwapValues(firstValue As Variant, secondValue As Variant) ' This sub switches the values
	Dim tmpValue As Variant ' Temp variable to store the first value
	tmpValue = firstValue ' put the first value into a temp variable
	firstValue = secondValue ' put the second value into the first
	secondValue = tmpValue ' and then put the first value, that stored in a temp variable, into the second
End Sub
If this code wasn't helpful and you still want to know how the bubble sort algorithm works so go to this links:
I hope this code was helpful if so please vote for me.
1) http://technology.niagarac.on.ca/courses/comp435/labs/bubblesort.html
2) http://www.cis.ufl.edu/~ddd/cis3020/summer-97/lectures/lec16/tsld042.htm
3) http://www.enm.maine.edu/Courses/C/SourceCode/BUBBLE.html
4) http://www-ee.eng.hawaii.edu/Courses/EE150/Book/chap10/subsection2.1.2.2.html
5)http://www.scism.sbu.ac.uk/law/Section5/chap2/s5c2p13.html
```

