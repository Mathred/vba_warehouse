VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DictionaryUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public weekNumber As Integer
Function FWeekRange(Optional ws As Worksheet) As Range
    Dim i As Integer
    If ws Is Nothing Then Set ws = ActiveWorkbook.Sheets("DPP")
    i = 4
    Do While (Not (FIfWeek(ws.Cells(i, 3)))) And i < 500
    i = i + 1
    Loop
    If weekNumber = 1 Then
        Set FWeekRange = ws.Range(ws.Cells(2, 2), ws.Cells(i - 1, FLastColumn(ws)))
    Else
        Set FWeekRange = ws.Range(ws.Cells(i, 2), ws.Cells(FLastRow(ws), FLastColumn(ws)))
    End If
End Function
Function FIfWeek(cel As Range) As Boolean
    If InStr(cel.NumberFormat, "Неделя №") > 0 Then
        FIfWeek = True
    Else: FIfWeek = False
    End If
End Function
Sub RecordsHeader(ws As Worksheet)
    ws.Cells(1, 1) = "Date"
    ws.Cells(1, 2) = "Line"
    ws.Cells(1, 3) = "Shifts"
    ws.Cells(1, 4) = "Product ID"
    ws.Cells(1, 5) = "Production, kg"
    ws.Cells(1, 6) = "Raw material, pallets"
    ws.Cells(1, 7) = "Line capacity, kg/h"
    ws.Cells(1, 8) = "Production, pallets"
End Sub
Function FLineRenamer(str As String) As String
    Select Case str
        Case "Resource : W70101_M046_001"
            FLineRenamer = "Линия 1"
        Case "Resource : W70105_M046_001"
            FLineRenamer = "Линия 5"
        Case "Resource : W70107_M046_001"
            FLineRenamer = "Линия 7"
        Case "Resource : W70108_M046_001"
            FLineRenamer = "Линия 8"
        Case "Resource : W70110_M046_001"
            FLineRenamer = "Линия 10"
        Case "Resource : W70111_M046_001"
            FLineRenamer = "Линия 11"
        Case "Resource : W70112_M046_001"
            FLineRenamer = "Линия 12"
        Case "Resource : W70113_M046_001"
            FLineRenamer = "Линия 13"
        Case "Resource : W70115_M046_001"
            FLineRenamer = "Линия 15"
        Case "Resource : W70116_M046_001"
            FLineRenamer = "Линия 16"
        Case "Resource : W70152_M046_001"
            FLineRenamer = "Линия 52"
        Case "Resource : W70103_M046_001"
            FLineRenamer = "Линия 3"
        Case "Resource : W70106_M046_001"
            FLineRenamer = "Линия 6"
        Case "Resource : W70114_M046_001"
            FLineRenamer = "Линия 14"
        Case "Resource : W70117_M046_001"
            FLineRenamer = "Линия 17"
        Case "Resource : W70262_M046_001"
            FLineRenamer = "Линия 56"
    End Select
End Function
Function FLineRange(str As String, Optional ws As Worksheet, Optional prevLineRange As Range) As Range
    If ws Is Nothing Then Set ws = ActiveWorkbook.Worksheets("DPP")
    Dim rng As Range
    Set rng = FWeekRange(ws)
    Dim i1 As Integer
    Dim i As Integer
    
    If prevLineRange Is Nothing Then
        i = rng.row
    Else
        i = prevLineRange.End(xlUp).row
    End If
    
    Do While InStr(ws.Cells(i, 2), str) = 0
    i = i + 1
    Loop
    i1 = i
    i = i + 1
    Do While (InStr(ws.Cells(i, 2), "Линия") Or InStr(ws.Cells(i, 2), "ПАЛЛЕТ ГП")) = 0
    i = i + 1
    Loop
    Set FLineRange = ws.Range(ws.Cells(i1, 2), ws.Cells(i - 1, FLastColumn(ws)))
End Function
Function FDateRange(dat As Date, Optional ws As Worksheet) As Range
    If ws Is Nothing Then Set ws = ActiveWorkbook.Worksheets("DPP")
    
    Dim i As Integer
    i = 1
    Dim dateRow As Integer: dateRow = FWeekRange(ws).row + 1
    Dim lastColumn As Integer: lastColumn = FLastColumn(ws)
    Do While ws.Cells(dateRow, i) <> dat Or i > lastColumn
        i = i + 1
    Loop
    Set FDateRange = ws.Range(ws.Cells(3, i), ws.Cells(FLastRow(ws), i))
End Function
Function FDPP2(product As Long, linerange As Range, daterange As Range, Optional ws As Worksheet) As Long
    
    If ws Is Nothing Then Set ws = ActiveWorkbook.Worksheets("DPP")
    Dim row As Integer: row = FProductRow(linerange, product, 1, ws)
    If IsEmpty(ws.Cells(row, daterange.Column).Value) Then
        row = FProductRow(linerange, product, row, ws)
        FDPP2 = ws.Cells(row, daterange.Column).Value
    Else
        FDPP2 = ws.Cells(row, daterange.Column).Value
    End If
End Function
Function FShiftRow(rng As Range, Optional ws As Worksheet) As Range
    If ws Is Nothing Then Set ws = ActiveWorkbook.Worksheets("DPP")
    Dim i As Integer
    i = 1
    Do While InStr(rng.Cells(i, 1), "кол-во") = 0
    i = i + 1
    Loop
    Set FShiftRow = rng.Range(ws.Cells(i, 1), ws.Cells(i, rng.Columns.Count))
End Function
Function FPalletsCount(prod As Long, dpp As Long) As Long
    FPalletsCount = 0
    Dim i As Integer
    Dim rng As Range
    Set rng = FProductRowPivot(prod)
    If (rng.row = 1) Then
        FPalletsCount = 0
    End If
    For i = 3 To rng.Columns.Count
        If rng.Cells(1, i) > 0 Then
            FPalletsCount = FPalletsCount + WorksheetFunction.RoundUp(dpp * rng.Cells(1, i), 0)
        End If
    Next i
End Function
Function FProductRow(rng As Range, prod As Long, start As Integer, Optional ws As Worksheet) As Integer
    If ws Is Nothing Then Set ws = ActiveWorkbook.Worksheets("DPP")
    Dim i As Integer
    If (start = 1) Then
        i = 1
    Else
        i = start - rng.row + 2
    End If
    Do While rng.Cells(i, 1).Value <> prod And rng.Rows.Count >= i
       i = i + 1
    Loop
    FProductRow = rng.row + i - 1
End Function
Function FCapacityCol(Optional ws As Worksheet) As Range
    If ws Is Nothing Then Set ws = ActiveWorkbook.Worksheets("DPP")
    Dim i As Integer
    i = Utils.FLastColumn(ws)
    Do While InStr(ws.Cells(2, i), "Kg / Hour") = 0
    i = i - 1
    Loop
    Set FCapacityCol = ws.Range(ws.Cells(1, i), ws.Cells(FLastRow(ws), i))
End Function
Function FNetCol(Optional ws As Worksheet) As Range
    If ws Is Nothing Then Set ws = ActiveWorkbook.Worksheets("DPP")
    Dim i As Integer
    i = Utils.FLastColumn(ws)
    Do While InStr(ws.Cells(2, i), "Вес нетто") = 0
    i = i - 1
    Loop
    Set FNetCol = ws.Range(ws.Cells(1, i), ws.Cells(FLastRow(ws), i))
End Function
Sub RecordsAutofit(ws As Worksheet)
Dim columnCounter As Integer
For columnCounter = 1 To Utils.FLastColumn(ws)
    ws.Cells(1, columnCounter).EntireColumn.AutoFit
Next columnCounter
End Sub
Private Function FProductRowPivot(prod As Long, Optional ws As Worksheet) As Range
    If ws Is Nothing Then Set ws = ActiveWorkbook.Worksheets("Pivot")
    Dim i As Integer: i = 4
    Do While ws.Cells(i, 1).Value <> prod And Utils.FLastRow(ws) >= i
        i = i + 1
    Loop
    If (i <= Utils.FLastRow(ws)) Then
        Set FProductRowPivot = ws.Range(ws.Cells(i, 1), ws.Cells(i, ws.Columns.Count))
    Else
        Set FProductRowPivot = ws.Range(ws.Cells(1, 1), ws.Cells(1, 1))
    End If
End Function
