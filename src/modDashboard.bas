Attribute VB_Name = "modDashboard"
Option Explicit
Sub RefreshPT()

Dim ws As Worksheet
Dim pt As PivotTable
Application.ScreenUpdating = False
For Each ws In ThisWorkbook.Worksheets
 For Each pt In ws.PivotTables
  pt.RefreshTable
 Next pt
Next ws

Application.ScreenUpdating = True
'MsgBox "PivotTables refreshed successfully.", vbInformation, "Dashboard Refresh"



End Sub

Sub UpdateDashboard()
Dim dashboardSheet, calcSheet As Worksheet
'Updating Dashboad Sheet
Set dashboardSheet = ThisWorkbook.Worksheets("Dashboard")
Set calcSheet = ThisWorkbook.Worksheets("Calculations")
Application.ScreenUpdating = False
dashboardSheet.Range("A6").Value = calcSheet.Range("B4").Value
dashboardSheet.Range("D6").Value = calcSheet.Range("B6").Value
dashboardSheet.Range("G6").Value = calcSheet.Range("B8").Value
dashboardSheet.Range("J6").Value = calcSheet.Range("B10").Value
dashboardSheet.Range("M6").Value = calcSheet.Range("B12").Value
'Standard Order
dashboardSheet.Range("B26").Value = calcSheet.Range("H4").Value
dashboardSheet.Range("C26").Value = calcSheet.Range("I4").Value
dashboardSheet.Range("D26").Value = calcSheet.Range("J4").Value
dashboardSheet.Range("E26").Value = calcSheet.Range("K4").Value
dashboardSheet.Range("F26").Value = calcSheet.Range("L4").Value

'Priority Order
dashboardSheet.Range("B27").Value = calcSheet.Range("H5").Value
dashboardSheet.Range("C27").Value = calcSheet.Range("I5").Value
dashboardSheet.Range("D27").Value = calcSheet.Range("J5").Value
dashboardSheet.Range("E27").Value = calcSheet.Range("K5").Value
dashboardSheet.Range("F27").Value = calcSheet.Range("L5").Value

'Bulk Order
dashboardSheet.Range("B28").Value = calcSheet.Range("H6").Value
dashboardSheet.Range("C28").Value = calcSheet.Range("I6").Value
dashboardSheet.Range("D28").Value = calcSheet.Range("J6").Value
dashboardSheet.Range("E28").Value = calcSheet.Range("K6").Value
dashboardSheet.Range("F28").Value = calcSheet.Range("L6").Value

With dashboardSheet
.Columns("C").NumberFormat = "0.00 ""min"""
.Columns("D:E").NumberFormat = "$0.00"
.Columns("F").NumberFormat = "0.0%"
.Columns("C:F").AutoFit
End With

dashboardSheet.Range("A39").Value = "Last Updated: " & Format(Now, "mm/dd/yyyy hh:mm AM/PM")
Application.ScreenUpdating = True

dashboardSheet.Range("A32").Value = "The highest average processing time occurs during " & _
                                    calcSheet.Range("B12").Value & ", making it the primary process bottleneck. "
                                    

End Sub

Sub ExportDashboardPDF()

Dim dashboardSheet As Worksheet
Dim savePath As Variant
Dim reportName As String
On Error GoTo ErrorHandler

Set dashboardSheet = ThisWorkbook.Worksheets("Dashboard")
reportName = "Warehouse_Operations_Report_" & Format(Date, "yyyy-mm-dd") & ".pdf"
savePath = Application.GetSaveAsFilename _
           (InitialFileName:=reportName, filefilter:="PDF files(*.pdf), *.pdf" _
            , Title:="Save Warehouse Operations Report")
If savePath = False Then
 MsgBox "PDF export cancelled.", vbInformation, "Export Cancelled"
 Exit Sub
End If

dashboardSheet.Range("A39").Value = "Last Updated: " & Format(Now, "mm/dd/yyyy hh:mm AM/PM")

dashboardSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=savePath, _
        Quality:=xlQualityStandard, _
        Includedocproperties:=True, _
        Ignoreprintareas:=False, _
        openafterpublish:=True
        
MsgBox "Dashboard exported successfully.", vbInformation, "Export Complete"
Exit Sub

ErrorHandler:
 MsgBox "The dashboard could not be reported." & vbCrLf & "Error:" & Err.Description, _
        vbCritical, "Export Error"
        
 
        
End Sub
