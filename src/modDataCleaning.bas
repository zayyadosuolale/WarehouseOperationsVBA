Attribute VB_Name = "modDataCleaning"
Option Explicit

Sub CleanData()

Dim rawSheet As Worksheet
Dim cleanSheet As Worksheet
Dim lastR, lastC, rowNum As Long
Dim cleanTable As ListObject
Dim tableRange As Range
Set rawSheet = ThisWorkbook.Worksheets("Raw_Data")
Set cleanSheet = ThisWorkbook.Worksheets("Clean_Data")

If WorksheetFunction.CountA(rawSheet.Cells) = 0 Then
 MsgBox "No raw data found. Please import a CSV file first.", _
        vbExclamation, "No Data "
        
 Exit Sub
 
End If

        
lastR = rawSheet.Cells(rawSheet.Rows.Count, 1).End(xlUp).Row
lastC = rawSheet.Cells(1, rawSheet.Columns.Count).End(xlToLeft).Column

If cleanSheet.ListObjects.Count > 0 Then
    cleanSheet.ListObjects(1).Unlist
End If

cleanSheet.Cells.Clear

rawSheet.Range(rawSheet.Cells(1, 1), rawSheet.Cells(lastR, lastC)) _
.Copy Destination:=cleanSheet.Range("A1")

cleanSheet.Range(cleanSheet.Cells(1, 1), cleanSheet.Cells(lastR, lastC)) _
.RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10), Header:=xlYes

lastR = cleanSheet.Cells(cleanSheet.Rows.Count, 1).End(xlUp).Row
'Creating clean sheet to make it easier for calculating metrics
For rowNum = 2 To lastR
    cleanSheet.Cells(rowNum, 1).Value = Trim(cleanSheet.Cells(rowNum, 1).Value)
    cleanSheet.Cells(rowNum, 3).Value = Trim(cleanSheet.Cells(rowNum, 3).Value)
    cleanSheet.Cells(rowNum, 4).Value = Trim(cleanSheet.Cells(rowNum, 4).Value)
    cleanSheet.Cells(rowNum, 10).Value = Trim(cleanSheet.Cells(rowNum, 10).Value)

   If cleanSheet.Cells(rowNum, 5).Value <= 0 Or _
    cleanSheet.Cells(rowNum, 6).Value <= 0 Or _
    cleanSheet.Cells(rowNum, 7).Value <= 0 Or _
    cleanSheet.Cells(rowNum, 8).Value <= 0 Or _
    cleanSheet.Cells(rowNum, 9).Value <= 0 Then
    
    cleanSheet.Rows(rowNum).Interior.ColorIndex = 6
    
   End If
Next rowNum

With cleanSheet
  .Rows(1).Font.Bold = True
  .Columns("B").NumberFormat = "mm/dd/yyyy"
  .Columns("E").NumberFormat = "0.0"
  .Columns("G").NumberFormat = "$0.00"
  .Columns("I").NumberFormat = "0.0"
  .Columns.AutoFit
  
  '.Activate
  '.Range("A1").Select
  
End With

'MsgBox (lastR - 1) & " records cleaned successfully. ", vbInformation, "Cleaning Complete"

cleanSheet.Range("K1").Value = "Labor_Hours"
cleanSheet.Range("L1").Value = "Labor_Cost"
cleanSheet.Range("M1").Value = "Units_Per_Hour"
cleanSheet.Range("N1").Value = "Cost_Per_Unit"
cleanSheet.Range("O1").Value = "Time_Variance"
cleanSheet.Range("P1").Value = "Target_Met"

Set tableRange = cleanSheet.Range("A1:P" & lastR)

Set cleanTable = cleanSheet.ListObjects.Add( _
    SourceType:=xlSrcRange, _
    Source:=tableRange, _
    XlListObjectHasHeaders:=xlYes)

cleanTable.Name = "tblCleanData"

End Sub
