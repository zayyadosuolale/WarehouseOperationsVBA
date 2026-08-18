Attribute VB_Name = "modCalculations"
Option Explicit
Sub CalculateM()

Dim cleanSheet As Worksheet
Dim lastR, rowNum As Long
Dim cleanTable As ListObject

Set cleanSheet = ThisWorkbook.Worksheets("Clean_data")
'Check to see whether the data is cleaned, but canceled out
If WorksheetFunction.CountA(cleanSheet.Cells) = 0 Then
 MsgBox "No cleaned data found. Please cleanData first. ", _
         vbExclamation, "No data "
  Exit Sub
End If
Application.ScreenUpdating = False

lastR = cleanSheet.Cells(cleanSheet.Rows.Count, 1).End(xlUp).Row
'Writing titles for the columns
With cleanSheet
  .Range("K1").Value = "Labor_Hours"
  .Range("L1").Value = "Labor_Cost"
  .Range("M1").Value = "Units_Per_Hour"
  .Range("N1").Value = "Cost_Per_Unit"
  .Range("O1").Value = "Time_Variance "
  .Range("P1").Value = "Target_Met"
'Calculate Metrics with Formulas
For rowNum = 2 To lastR
 .Cells(rowNum, 11).Value = (.Cells(rowNum, 5).Value / 60) * (.Cells(rowNum, 6).Value)
 .Cells(rowNum, 12).Value = (.Cells(rowNum, 11).Value) * (.Cells(rowNum, 7).Value)
 .Cells(rowNum, 13).Value = (.Cells(rowNum, 8).Value) / (.Cells(rowNum, 5).Value / 60)
 .Cells(rowNum, 14).Value = (.Cells(rowNum, 12).Value) / (.Cells(rowNum, 8).Value)
 .Cells(rowNum, 15).Value = (.Cells(rowNum, 5).Value) - (.Cells(rowNum, 9).Value)
 
 If .Cells(rowNum, 5).Value <= .Cells(rowNum, 9).Value Then
  .Cells(rowNum, 16).Value = "Yes"
 Else
  .Cells(rowNum, 16).Value = "No"
 End If
 
  If .Cells(rowNum, 16).Value = "Yes" Then
   .Cells(rowNum, 16).Interior.Color = RGB(198, 239, 206)
  Else
  .Cells(rowNum, 16).Interior.Color = RGB(255, 199, 206)
  End If
  If .Cells(rowNum, 15).Value > 0 Then
   .Cells(rowNum, 15).Interior.Color = RGB(255, 199, 206)
  Else
  .Cells(rowNum, 15).Interior.Color = RGB(198, 239, 206)
  End If
  
 
 Next rowNum

  .Columns("K").NumberFormat = "0.00"
  .Columns("L").NumberFormat = "$0.00"
  .Columns("M").NumberFormat = "0.00"
  .Columns("N").NumberFormat = "$0.00"
  .Columns("O").NumberFormat = "0.00"
  .Columns.AutoFit
  .Range("k1:p1").Font.Bold = True
  
 '.Activate
' .Range("a1").Select
  
End With

Application.ScreenUpdating = True
'MsgBox (lastR - 1) & " records analyzed successfully. ", _
   '  vbInformation, "Metrics Complete"
Set cleanTable = cleanSheet.ListObjects("tblCleanData")
cleanTable.Resize cleanSheet.Range("A1:P" & lastR)
End Sub
