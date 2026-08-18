Attribute VB_Name = "modMain"
Option Explicit
Sub RunFullAnalysis()

On Error GoTo ErrorHandler
Application.ScreenUpdating = False
Application.EnableEvents = False

'Step 1: Import CSV
If importCSV = False Then
 Application.screnupdating = True
 Application.EnableEvents = True
Exit Sub
End If

'Step 2: Clean imported Data
CleanData
'Step 3: Calculate row-level metrics
CalculateM
'Step 4: Build management KPIs
BuildKPIs
'Step 5: Refresh PivotTables and Charts
RefreshPT
'Step 6: Update Dashboard
UpdateDashboard

'Show finished dashboard
ThisWorkbook.Worksheets("Dashboard").Activate
ThisWorkbook.Worksheets("Dashboard").Range("A1").Select
Application.ScreenUpdating = True
Application.EnableEvents = True
MsgBox "Warehouse analysis completed successfully.", vbInformation, "Analysis Complete"
Exit Sub

ErrorHandler:
 Application.ScreenUpdating = True
 Application.EnableEvents = True
 
 MsgBox "The analysis could not be completed. " & vbCrLf & "Error:" _
        & Err.Description, vbCritical, "Analaysis Error"
 
End Sub

Sub importDataOnly()
 If importCSV() Then
  MsgBox "Data imported successfully.", vbInformation, "Import Complete"
 End If
End Sub

Sub RefreshDashboardOnly()
On Error GoTo ErrorHandler

Application.ScreenUpdating = False

CalculateM
BuildKPIs
RefreshPT
UpdateDashboard

ThisWorkbook.Worksheets("Dashboard").Activate
Application.ScreenUpdating = True

MsgBox "Dashboard refreshed successfully.", vbInformation, "Refresh Complete "

Exit Sub

ErrorHandler:

 Application.ScreenUpdating = True
 MsgBox "Dashbaord refresh failed." & vbCrLf & "Error:" & Err.Description, vbCritical, "Refresh Error"


End Sub
