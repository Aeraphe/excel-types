'/*
'A collection of all the ListObject objects on a worksheet. Each ListObject object represents a table on the worksheet.
'*/
Public Class ListObject()

'/*
'Deletes the ListObject object and clears the cell data from the worksheet.
'*/
Public Function Delete()

End Function

'/*
'Exports a ListObject object to Visio.
'*/
Public Function ExportToVisio()

End Function

'/*
'Publishes the ListObject object to a server that is running Microsoft SharePoint Foundation.
'The Target parameter contains an array of String elements, as described in the following table.
'0 : URL of SharePoint server
'1 : ListName (Display Name)
'2 : Description of the list. Optional.
'
'@param {Variant} Target
'@param {Boolean} LinkSource
'@return String
'*/
Public Function Publish(Target, LinkSource) As String

End Function



End Class