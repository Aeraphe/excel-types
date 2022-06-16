'/*
'A collection of all the ListObject objects on a worksheet. Each ListObject object represents a table on the worksheet.
'*/
Public Class ListObjects()

'/*
'Creates a new list object.
'
'@param {XlListObjectSourceType} SourceType:[Optional]
'@param {Variant} Source:[Optional]
'@param {Boolean} LinkSource:[Optional]
'@param {Variant} XlListObjectHasHeaders:[Optional]
'@param {Variant} Destination:[Optional]
'@param {String} TableStyleName:[Optional]
'@return ListObject
'*/
Public Function Add(SourceType, Source, LinkSource, XlListObjectHasHeaders, Destination, TableStyleName) As ListObject

End Function

'/*
'When used without an object qualifier, this property returns an Application object that represents the Microsoft Excel application.
'When used with an object qualifier, this property returns an Application object that represents the creator of the specified object
'(you can use this property with an OLE Automation object to return the application of that object). Read-only.
'
'@type {Object.<Application>}
'*/
Public Property Application As Application

'/*
'Returns an Integer value that represents the number of objects in the collection.
'
'@type {Integer}
'*/
Public Property Count As Integer

'/*
'Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
'
'@type {Long}
'*/
Public Property Creator As Long

'/*
'Returns a single object from a collection.
'
'@param {Variant} Index
'@type {Object.<WorkSheet>}
'*/
Public Property Item(Index) As ListObject

'/*
'Returns the parent object for the specified object. Read-only.
'
'@type {Object.<WorkSheet>}
'*/
Public Property Parent As WorkSheet

End Class
