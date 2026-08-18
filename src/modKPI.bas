Attribute VB_Name = "modKPI"
Option Explicit

Sub BuildKPIs()
Dim cleanSheet As Worksheet
Dim calcSheet As Worksheet
Dim lastR, rowNum As Long
Dim orderUnits, orderTimes, orderCosts, stepTotal, stepCount As Object
Dim orderID, processStep As String
Dim duration, units, laborCost As Double
Dim totalOrders As Long
Dim totalUnits, totalOT, totalLC As Double
Dim key As Variant
Dim avgOT, avgCPO, avgCPU As Double
Dim targetsMet, totalTargets, outputRow, pathRow As Long
Dim targetRate As Double
Dim bottleNS As String
Dim bottleNT, avgST As Double
Dim orderPaths, pathOrders, pathUnits, pathTimes, pathCosts, pathsTM, pathsTT As Object
Dim processPath, currentPath As String


Set cleanSheet = ThisWorkbook.Worksheets("Clean_data")
Set calcSheet = ThisWorkbook.Worksheets("Calculations")

If WorksheetFunction.CountA(cleanSheet.Cells) = 0 Then
 MsgBox "No analyzed data found. Please run CalculateMetrics first. ", _
         vbExclamation, "No data "
  Exit Sub
End If

Application.ScreenUpdating = False
calcSheet.Cells.Clear
Set orderUnits = CreateObject("Scripting.Dictionary")
Set orderTimes = CreateObject("Scripting.Dictionary")
Set orderCosts = CreateObject("Scripting.Dictionary")
Set stepTotal = CreateObject("Scripting.Dictionary")
Set stepCount = CreateObject("Scripting.Dictionary")
Set orderPaths = CreateObject("Scripting.Dictionary")
Set pathOrders = CreateObject("Scripting.Dictionary")
Set pathUnits = CreateObject("Scripting.Dictionary")
Set pathTimes = CreateObject("Scripting.Dictionary")
Set pathCosts = CreateObject("Scripting.Dictionary")
Set pathsTM = CreateObject("Scripting.Dictionary")
Set pathsTT = CreateObject("Scripting.Dictionary")

lastR = cleanSheet.Cells(cleanSheet.Rows.Count, 1).End(xlUp).Row

With cleanSheet

 For rowNum = 2 To lastR
  orderID = .Cells(rowNum, 1).Value
  processPath = .Cells(rowNum, 3).Value
  processStep = .Cells(rowNum, 4).Value
  duration = .Cells(rowNum, 5).Value
  units = .Cells(rowNum, 8).Value
  laborCost = .Cells(rowNum, 12).Value

'Order Units
  If Not orderUnits.exists(orderID) Then
   orderUnits.Add orderID, units
  End If
'Order Times
  If orderTimes.exists(orderID) Then
   orderTimes(orderID) = orderTimes(orderID) + duration
  Else
   orderTimes.Add orderID, duration
  End If
'Order Costs
  If orderCosts.exists(orderID) Then
   orderCosts(orderID) = orderCosts(orderID) + laborCost
  Else
   orderCosts.Add orderID, laborCost
  End If
'Stepchek
  If stepTotal.exists(processStep) Then
    stepTotal(processStep) = stepTotal(processStep) + duration
    stepCount(processStep) = stepCount(processStep) + 1
  Else
   stepTotal.Add processStep, duration
   stepCount.Add processStep, 1
  End If
  'Pathcreation
  If Not orderPaths.exists(orderID) Then
   orderPaths.Add orderID, processPath
  End If
  
  totalTargets = totalTargets + 1
  If .Cells(rowNum, 16).Value = "Yes" Then
   targetsMet = targetsMet + 1
  End If
  
  If pathsTT.exists(processPath) Then
   pathsTT(processPath) = pathsTT(processPath) + 1
  Else
   pathsTT.Add processPath, 1
  End If
  
  If .Cells(rowNum, 16).Value = "Yes" Then
   If pathsTM.exists(processPath) Then
    pathsTM(processPath) = pathsTM(processPath) + 1
   Else
    pathsTM.Add processPath, 1
    
   End If
   
  End If
     
 Next rowNum
End With


totalOrders = orderUnits.Count
For Each key In orderUnits.keys
 totalUnits = totalUnits + orderUnits(key)
 totalOT = totalOT + orderTimes(key)
 totalLC = totalLC + orderCosts(key)
Next key

avgOT = totalOT / totalOrders
avgCPO = totalLC / totalOrders
avgCPU = totalLC / totalUnits
targetRate = targetsMet / totalTargets

For Each key In orderUnits.keys
 currentPath = orderPaths(key)
 'Orders by path
 If pathOrders.exists(currentPath) Then
  pathOrders(currentPath) = pathOrders(currentPath) + 1
 Else
  pathOrders.Add currentPath, 1
 End If
 'Units by Path
 If pathUnits.exists(currentPath) Then
  pathUnits(currentPath) = pathUnits(currentPath) + orderUnits(key)
 Else
  pathUnits.Add currentPath, orderUnits(key)
 End If
 'Time by path
 If pathTimes.exists(currentPath) Then
  pathTimes(currentPath) = pathTimes(currentPath) + orderTimes(key)
 Else
  pathTimes.Add currentPath, orderTimes(key)
 End If
 'labor cost by path
 If pathCosts.exists(currentPath) Then
  pathCosts(currentPath) = pathCosts(currentPath) + orderCosts(key)
 Else
  pathCosts.Add currentPath, orderCosts(key)
 End If
 
Next key



With calcSheet
 .Range("A1").Value = "Warehouse Operations KPI Summary"
 .Range("A3").Value = "Metric"
 .Range("B3").Value = "Value"
 .Range("A4").Value = "Total Orders"
 .Range("A5").Value = "Total Units"
 .Range("A6").Value = "Average Order Time"
 .Range("A7").Value = "Total Labor Cost"
 .Range("A8").Value = "Average Cost Per Order"
 .Range("A9").Value = "Average Cost Per unit"
 .Range("A10").Value = "Target Achievement %"
 .Range("B4").Value = totalOrders
 .Range("B5").Value = totalUnits
 .Range("B6").Value = avgOT
 .Range("B7").Value = totalLC
 .Range("B8").Value = avgCPO
 .Range("B9").Value = avgCPU
 .Range("B10").Value = targetRate
 
 
 With .Range("A1")
   With .Font
   .Bold = True
   .Size = 16
   End With
 End With
 
 .Range("A3:B3").Font.Bold = True
 .Range("B6").NumberFormat = "0.00 ""min"""
 .Range("B7:B9").NumberFormat = "$0.00"
 .Range("B10").NumberFormat = "0.0%"
 .Columns("A:B").AutoFit
 
 .Range("D3").Value = "Process Step"
 .Range("E3").Value = "Avg Duration"
 outputRow = 4
 For Each key In stepTotal.keys
  .Cells(outputRow, 4).Value = key
  .Cells(outputRow, 5).Value = stepTotal(key) / stepCount(key)
  outputRow = outputRow + 1
  Next key
  
  .Range("D3 : E3").Font.Bold = True
  .Columns("E").NumberFormat = "0.00""min"""
  .Columns("D:E").AutoFit
  
  For Each key In stepTotal.keys
   avgST = stepTotal(key) / stepCount(key)
   If avgST > bottleNT Then
    bottleNT = avgST
    bottleNS = key
   End If
  Next key
  
   .Range("A12").Value = "Bottleneck Step"
   .Range("B12").Value = bottleNS
   .Range("A13").Value = "Bottleneck Avg Duration"
   .Range("B13").Value = bottleNT
   .Range("B13").NumberFormat = "0.00""min"""
   
   '.Activate
   '.Range("A1").Select
End With

pathRow = 4
With calcSheet
.Range("G3").Value = "Process Path"
.Range("H3").Value = "Orders"
.Range("I3").Value = "Avg Order Time"
.Range("J3").Value = "Avg Labor Cost"
.Range("K3").Value = "Avg Cost Per Unit"
.Range("L3").Value = "Target %"
 For Each key In pathOrders.keys
  .Cells(pathRow, 7).Value = key
  .Cells(pathRow, 8).Value = pathOrders(key)
  .Cells(pathRow, 9).Value = pathTimes(key) / pathOrders(key)
  .Cells(pathRow, 10).Value = pathCosts(key) / pathOrders(key)
  .Cells(pathRow, 11).Value = pathCosts(key) / pathUnits(key)
  If pathsTM.exists(key) Then
   .Cells(pathRow, 12).Value = pathsTM(key) / pathsTT(key)
  Else
   .Cells(pathRow, 12).Value = 0
  End If
 
  pathRow = pathRow + 1
 
 Next key
 .Range("G3:L3").Font.Bold = True
 .Columns("I").NumberFormat = "0.00 ""min"""
 .Columns("J:K").NumberFormat = "$0.00"
 .Columns("L").NumberFormat = "0.0%"
 .Columns("G:L").AutoFit
 
 
End With


Application.ScreenUpdating = True
'MsgBox "KPI summary created successfully.", vbInformation, "KPI Analysis Complete"


End Sub
