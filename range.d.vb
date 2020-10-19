'/*
'Represents a cell, a row, a column, a selection of cells containing one or more contiguous blocks of cells, or a 3D range.
'
'
'*/
Public Class Range()

'/*
'(Range)
'
'Activates a single cell, which must be inside the current selection. 
'To select a range of cells, use the Select method., 
'
'Example
'
'This example selects cells A1:C3 on Sheet1 and then makes cell B2 the active cell.
'
' Worksheets("Sheet1").Activate 
' Range("A1:C3").Select 
' Range("B2").Activate
'
'*/    
Public Sub Activate() 

End Sub


'/*
'Adds a comment to the range.
'
'Example:
'
'Worksheets(1).Range("E5").AddComment "Current Sales"
'
'@param {String} text
'*/
Public  Sub AddComment(text As String)

End Sub


End Class
