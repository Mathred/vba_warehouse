VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TablesUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub TablesFormat(ws As Worksheet)
    Dim vCell As Range
    Dim VTime As Date
    Dim rng As Range
    Dim PalletsSumRGB As Long: PalletsSumRGB = RGB(200, 225, 255)
    
    Set rng = ws.Range(ws.Cells(1, 2), ws.Cells(1, 49))
    rng.ColumnWidth = 2.2
    rng.NumberFormat = "hh:mm"
    rng.Orientation = xlUpward
    
    VTime = "08:00"
    
    For Each vCell In rng
    
    vCell.Value = VTime
    VTime = VTime + "00:30"
    
    Next vCell
    
    rng.Rows.AutoFit
    
    With ws
    .Cells(2, 1).Value = "Линия 1"
    .Cells(3, 1).Value = "Линия 5"
    .Cells(4, 1).Value = "Линия 7"
    .Cells(5, 1).Value = "Линия 8"
    .Cells(6, 1).Value = "Линия 10"
    .Cells(7, 1).Value = "Линия 11"
    .Cells(8, 1).Value = "Линия 12"
    .Cells(9, 1).Value = "Линия 13"
    .Cells(10, 1).Value = "Линия 15"
    .Cells(11, 1).Value = "Линия 16"
    .Cells(12, 1).Value = "Линия 52"
    '.Cells(13, 2).Value = "Первая смена BAP"
    '.Cells(13, 27).Value = "Вторая смена BAP"
    '.Cells(13, 43).Value = "ИТОГО за сутки BAP"
    '.Cells(14, 1).Value = "Паллетов FG BAP"
    '.Cells(15, 1).Value = "Паллетов RM BAP"
    .Cells(14, 1).Value = "Линия 3"
    .Cells(15, 1).Value = "Линия 6"
    .Cells(16, 1).Value = "Линия 14"
    .Cells(17, 1).Value = "Линия 17"
    .Cells(18, 1).Value = "Линия 56"
    '.Cells(22, 2).Value = "Первая смена NDC"
    '.Cells(22, 27).Value = "Вторая смена NDC"
    '.Cells(22, 43).Value = "ИТОГО за сутки NDC"
    '.Cells(23, 1).Value = "Паллетов FG NDC"
    '.Cells(24, 1).Value = "Паллетов RM NDC"
    '.Cells(26, 2).Value = "Первая смена"
    '.Cells(26, 27).Value = "Вторая смена"
    '.Cells(26, 43).Value = "ИТОГО за сутки"
    '.Cells(27, 1).Value = "Паллетов RM ИТОГО"
    '.Cells(28, 1).Value = "Паллетов FG ИТОГО"
    
    Application.DisplayAlerts = False
    
    'Set rng = .Range(.Cells(13, 2), .Cells(13, 8))
    'rng.Merge
    'rng.Interior.color = PalletsSumRGB
    'Set rng = .Range(.Cells(13, 27), .Cells(13, 33))
    'rng.Merge
    'rng.Interior.color = PalletsSumRGB
    'Set rng = .Range(.Cells(13, 43), .Cells(13, 49))
    'rng.Merge
    'rng.Interior.color = PalletsSumRGB
    
    'Set rng = .Range(.Cells(14, 2), .Cells(14, 8))
    'rng.Merge
    'rng.Interior.color = PalletsSumRGB
    'Set rng = .Range(.Cells(14, 27), .Cells(14, 33))
    'rng.Merge
    'rng.Interior.color = PalletsSumRGB
    'Set rng = .Range(.Cells(14, 43), .Cells(14, 49))
    'rng.Merge
    'rng.Interior.color = PalletsSumRGB
    
    'Set rng = .Range(.Cells(15, 2), .Cells(15, 8))
    'rng.Merge
    'rng.Interior.color = PalletsSumRGB
    'Set rng = .Range(.Cells(15, 27), .Cells(15, 33))
    'rng.Merge
    'rng.Interior.color = PalletsSumRGB
    'Set rng = .Range(.Cells(15, 43), .Cells(15, 49))
    'rng.Merge
    'rng.Interior.color = PalletsSumRGB
    
    'Set rng = .Range(.Cells(22, 2), .Cells(22, 8))
    'rng.Merge
    'rng.Interior.color = PalletsSumRGB
    'Set rng = .Range(.Cells(22, 27), .Cells(22, 33))
    'rng.Merge
    'rng.Interior.color = PalletsSumRGB
    'Set rng = .Range(.Cells(22, 43), .Cells(22, 49))
    'rng.Merge
    'rng.Interior.color = PalletsSumRGB
    
    
    'Set rng = .Range(.Cells(23, 2), .Cells(23, 8))
    'rng.Merge
    'rng.Interior.color = PalletsSumRGB
    'Set rng = .Range(.Cells(23, 27), .Cells(23, 33))
    'rng.Merge
    'rng.Interior.color = PalletsSumRGB
    'Set rng = .Range(.Cells(23, 43), .Cells(23, 49))
    'rng.Merge
    'rng.Interior.color = PalletsSumRGB
    
    'Set rng = .Range(.Cells(24, 2), .Cells(24, 8))
    'rng.Merge
    'rng.Interior.color = PalletsSumRGB
    'Set rng = .Range(.Cells(24, 27), .Cells(24, 33))
    'rng.Merge
    'rng.Interior.color = PalletsSumRGB
    'Set rng = .Range(.Cells(24, 43), .Cells(24, 49))
    'rng.Merge
    'rng.Interior.color = PalletsSumRGB
    
    'Set rng = .Range(.Cells(26, 2), .Cells(26, 8))
    'rng.Merge
    'rng.Interior.color = PalletsSumRGB
    'Set rng = .Range(.Cells(26, 27), .Cells(26, 33))
    'rng.Merge
    'rng.Interior.color = PalletsSumRGB
    'Set rng = .Range(.Cells(26, 43), .Cells(26, 49))
    'rng.Merge
    'rng.Interior.color = PalletsSumRGB
    
    'Set rng = .Range(.Cells(27, 2), .Cells(27, 8))
    'rng.Merge
    'rng.Interior.color = PalletsSumRGB
    'Set rng = .Range(.Cells(27, 27), .Cells(27, 33))
    'rng.Merge
    'rng.Interior.color = PalletsSumRGB
    'Set rng = .Range(.Cells(27, 43), .Cells(27, 49))
    'rng.Merge
    'rng.Interior.color = PalletsSumRGB
    
    'Set rng = .Range(.Cells(28, 2), .Cells(28, 8))
    'rng.Merge
    'rng.Interior.color = PalletsSumRGB
    'Set rng = .Range(.Cells(28, 27), .Cells(28, 33))
    'rng.Merge
    'rng.Interior.color = PalletsSumRGB
    'Set rng = .Range(.Cells(28, 43), .Cells(28, 49))
    'rng.Merge
    'rng.Interior.color = PalletsSumRGB
    
    Application.DisplayAlerts = True
    End With
    
    ws.Cells(1, 1).EntireColumn.AutoFit
    
    Call SetTablesRangeBorder(ws.UsedRange)
End Sub
Private Sub SetTablesRangeBorder(poRng As Range)
    If Not poRng Is Nothing Then
        poRng.Borders(xlDiagonalDown).LineStyle = xlNone
        poRng.Borders(xlDiagonalUp).LineStyle = xlNone
        poRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
        poRng.Borders(xlEdgeTop).LineStyle = xlContinuous
        poRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
        poRng.Borders(xlEdgeRight).LineStyle = xlContinuous
        poRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        poRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End If
End Sub
Sub CreateSimpleRecord(rng As Range, record As String)
    Dim vCell As Range
        For Each vCell In rng
            vCell.Value = record
        Next vCell
End Sub
Function LineMatcher(lineName As String) As Integer
    Select Case lineName
        Case "Линия 1"
            LineMatcher = 2
        Case "Линия 5"
            LineMatcher = 3
        Case "Линия 7"
            LineMatcher = 4
        Case "Линия 8"
            LineMatcher = 5
        Case "Линия 10"
            LineMatcher = 6
        Case "Линия 11"
            LineMatcher = 7
        Case "Линия 12"
            LineMatcher = 8
        Case "Линия 13"
            LineMatcher = 9
        Case "Линия 15"
            LineMatcher = 10
        Case "Линия 16"
            LineMatcher = 11
        Case "Линия 52"
            LineMatcher = 12
        Case "Линия 3"
            LineMatcher = 14
        Case "Линия 6"
            LineMatcher = 15
        Case "Линия 14"
            LineMatcher = 16
        Case "Линия 17"
            LineMatcher = 17
        Case "Линия 56"
            LineMatcher = 18
    End Select
End Function
Sub CreateIDRecord(rng As Range, id As String)
    Application.DisplayAlerts = False
    
    Dim vCell As Range
    For Each vCell In rng
        vCell.Value = id
    Next vCell
    'rng.Cells(1, 1).Interior.color = color
    
    Application.DisplayAlerts = True

End Sub
Sub CreateFGRecord(rng As Range, amount As Double, capacity As Double)
    Dim vCell As Range
    Dim interQuantity As Double
    Application.DisplayAlerts = False
    interQuantity = 0
    If rng.Columns.Count > 1 Then
        For Each vCell In rng
            If interQuantity + capacity / 2 < amount Then
                vCell.Value = capacity / 2
                interQuantity = interQuantity + vCell.Value
                'Debug.Print interQuantity
            Else
                vCell.Value = amount - interQuantity
            End If
        Next vCell
    Else
        rng.Cells(1, 1).Value = amount
    End If
    
    
    'For Each vCell In rng
    '    vCell.Value = id
    'Next vCell
    'rng.Cells(1, 1).Interior.color = color
    
    Application.DisplayAlerts = True

End Sub
Sub MergeTables(targetName As String)
    Application.DisplayAlerts = False
    
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim wsTarget As Worksheet
    Dim wsDate As Worksheet
    Dim TablesUtils As New TablesUtils
    Dim columnCounter As Integer
    Dim rowCounter As Integer
    Dim dateTableColumn As Integer
    
    If Utils.FSheetExists(targetName, wb) Then
        wb.Sheets(targetName).Delete
        Utils.SCreateWS (targetName)
    Else
        Utils.SCreateWS (targetName)
    End If
    
    Set wsTarget = wb.Sheets(targetName)
    'Call TablesUtils.TablesFormat(wsTarget)
    columnCounter = 1
    For Each wsDate In wb.Sheets
        If IsDate(wsDate.Name) Then
            For dateTableColumn = 2 To 33
                wsTarget.Cells(1, dateTableColumn + columnCounter - 1).Value = wsDate.Name & " " & Format(wsDate.Cells(1, dateTableColumn), "hh:mm")
            Next dateTableColumn
            For dateTableColumn = 34 To 49
                wsTarget.Cells(1, dateTableColumn + columnCounter - 1).Value = Format(CDate(wsDate.Name) + 1, "dd.mm.yyyy") & " " & Format(wsDate.Cells(1, dateTableColumn), "hh:mm")
            Next dateTableColumn
            For rowCounter = 2 To 18
                For dateTableColumn = 2 To 49
                    wsTarget.Cells(rowCounter, dateTableColumn + columnCounter - 1).Value = Format(wsDate.Cells(rowCounter, dateTableColumn).Value, "#")
                Next dateTableColumn
            Next rowCounter
            
            
            
            columnCounter = columnCounter + 48
        End If
    Next wsDate
    
    Call TablesUtils.FGIDTablesFormat(wsTarget)
    
    Application.DisplayAlerts = True
End Sub

Sub FGIDTablesFormat(ws As Worksheet)
    
    Application.DisplayAlerts = False
    
    Dim vCell As Range
    Dim VTime As Date
    Dim rng As Range
    
    Set rng = ws.Range(ws.Cells(1, 2), ws.Cells(Utils.FLastRow(ws), Utils.FLastColumn(ws) + 1))
    rng.ColumnWidth = 2.2
    rng.Orientation = xlUpward
    
    rng.Rows.AutoFit
    
    With ws
        .Cells(2, 1).Value = "Линия 1"
        .Cells(3, 1).Value = "Линия 5"
        .Cells(4, 1).Value = "Линия 7"
        .Cells(5, 1).Value = "Линия 8"
        .Cells(6, 1).Value = "Линия 10"
        .Cells(7, 1).Value = "Линия 11"
        .Cells(8, 1).Value = "Линия 12"
        .Cells(9, 1).Value = "Линия 13"
        .Cells(10, 1).Value = "Линия 15"
        .Cells(11, 1).Value = "Линия 16"
        .Cells(12, 1).Value = "Линия 52"
        .Cells(14, 1).Value = "Линия 3"
        .Cells(15, 1).Value = "Линия 6"
        .Cells(16, 1).Value = "Линия 14"
        .Cells(17, 1).Value = "Линия 17"
        .Cells(18, 1).Value = "Линия 56"
        
    End With
    
    ws.Cells(1, 1).EntireColumn.AutoFit
    
    'Call SetTablesRangeBorder(ws.UsedRang)
    Application.DisplayAlerts = True
End Sub

