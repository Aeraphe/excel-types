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

Public  Sub AddCommentThreaded()

End Sub

Public  Sub AdvancedFilter()

End Sub

Public  Sub AllocateChanges()

End Sub

Public  Sub ApplyName()

End Sub

Public  Sub ApplyOutLineStyles()

End Sub

Public  Sub AutoComplete()

End Sub

Public  Sub AutoFill()

End Sub

Public  Sub AutoFilter()

End Sub

Public  Sub AutioFit()

End Sub

Public  Sub AutoOutline()

End Sub

Public  Sub BorderAround()

End Sub

Public  Sub Calculate()

End Sub

Public  Sub CalculateRowMajorOrder()

End Sub

Public  Sub CheckSpelling()

End Sub

Public  Sub Clear()

End Sub

Public  Sub ClearComments()

End Sub

Public  Sub ClearContents()

End Sub

Public  Sub ClearFormats()

End Sub

Public  Sub ClearHyperlinks()

End Sub

Public  Sub ClearNotes()

End Sub

Public  Sub ClearOutline()

End Sub

Public  Sub ColumnDifferences()

End Sub

Public  Sub Consolidate()

End Sub

Public  Sub ConvertToLinkedDataType()

End Sub

Public  Sub Copy()

End Sub

Public  Sub CopyFromRecordset()

End Sub

Public  Sub CopyPicture()

End Sub

Public  Sub CreateNames()

End Sub

Public  Sub Cut()

End Sub

End Class
