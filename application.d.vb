'/*
'Represents the entire Microsoft Excel application.
'
'
'*/
Public Class Application()

'/*
'Returns a Range object that represents the active cell in the active window 
'(the window on top) or in the specified window. If the window isn't displaying 
'a worksheet, this property fails. Read-only.
'
'@type {Object.<Range>}
'
'*/
Public Property ActiveCell As Range

'/*
'Returns a Chart object that represents the active chart (either an embedded chart or a chart sheet). 
'An embedded chart is considered active when it's either selected or activated. When no chart is active, 
'this property returns Nothing.
'
'Example: 
'ActiveChart.HasLegend = True
'
'@type {Object.<Chart>}
'*/
Public Property ActiveChart As Chart

'/*
'Returns a Workbooks collection that represents all the open workbooks. Read-only.
'
'@type {Object.<Collection>} Workbooks Collection
'*/
Public Property Workbooks As Workbooks

'/*
'Activates a Microsoft application. If the application is already running, 
'this method activates the running application. 
'If the application isn't running, this method starts a new instance of the application.
'
'Example: (This example starts and activates Word.)
'
'Application.ActivateMicrosoftApp xlMicrosoftWord
'
'@param {XlMSApplication}  index 
'*/    
Public Sub ActivateMicrosoftApp( index As XlMSApplication) 

End Sub

'/*
'An event occurs when all pending refresh activity (both synchronous and asynchronous) 
'and all of the resultant calculation activities have been completed.
'
'*/
Public Event AfterCalculate()

'/*
'Occurs when a new workbook is created.
'
'Example:
'
'Private Sub App_NewWorkbook(ByVal Wb As Workbook) 
'Application.Windows.Arrange xlArrangeStyleTiled End Sub
'   
'@param {Workbook} Wb
'*/
Public Event NewWorkbook(ByVal Wb As Workbook) 


End Class