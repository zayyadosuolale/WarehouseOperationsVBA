Attribute VB_Name = "modDataImport"
Option Explicit

Function importCSV() As Boolean
importCSV = False
'Declaring Variables
  Dim filePath As Variant
  Dim rawSheet As Worksheet
  Dim csvWorkbook As Workbook
  Dim csvSheet As Worksheet
  Dim lastR As Long
  Dim lastC As Long


Set rawSheet = ThisWorkbook.Worksheets("Raw_Data")

filePath = Application.GetOpenFilename( _
  filefilter:="CSV Files (*.csv), *.csv", _
  Title:="Select Warehouse Order Data")
  

If filePath = False Then
 MsgBox "Import cancelled.", vbInformation, "Warehouse Dashboard"
 Exit Function
End If

MsgBox "File selected Successfully.", vbInformation, "Warehouse Dashboard"

Application.ScreenUpdating = False

rawSheet.Cells.Clear

Set csvWorkbook = Workbooks.Open(filePath)
Set csvSheet = csvWorkbook.Worksheets(1)

lastR = csvSheet.Cells(csvSheet.Rows.Count, 1).End(xlUp).Row
lastC = csvSheet.Cells(1, csvSheet.Columns.Count).End(xlToLeft).Column

csvSheet.Range(csvSheet.Cells(1, 1), csvSheet.Cells(lastR, lastC)).Copy Destination _
:=rawSheet.Range("A1")

csvWorkbook.Close savechanges:=False

rawSheet.Columns.AutoFit
rawSheet.Rows(1).Font.Bold = True

'rawSheet.Activate
'rawSheet.Range("A1").Select

Application.ScreenUpdating = True

'MsgBox (lastR - 1) & " process records imported sucessfully.", _
'vbInformation, "Import Complete"


importCSV = True



End Function
