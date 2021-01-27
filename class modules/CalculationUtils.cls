VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CalculationUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub SDeleteTimeTableSheets()
    Dim i As Integer
    Application.DisplayAlerts = False
    For i = Sheets.Count To 1 Step -1
        If (IsDate(Sheets(i).Name)) Then
            Sheets(i).Delete
        End If
    Next i
    Application.DisplayAlerts = True
End Sub
Sub STableTemplate(ws As Worksheet)
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
    .Cells(13, 2).Value = "Первая смена BAP"
    .Cells(13, 27).Value = "Вторая смена BAP"
    .Cells(13, 43).Value = "ИТОГО за сутки BAP"
    .Cells(14, 1).Value = "Паллетов FG BAP"
    .Cells(15, 1).Value = "Паллетов RM BAP"
    .Cells(17, 1).Value = "Линия 3"
    .Cells(18, 1).Value = "Линия 6"
    .Cells(19, 1).Value = "Линия 14"
    .Cells(20, 1).Value = "Линия 17"
    .Cells(21, 1).Value = "Линия 56"
    .Cells(22, 2).Value = "Первая смена NDC"
    .Cells(22, 27).Value = "Вторая смена NDC"
    .Cells(22, 43).Value = "ИТОГО за сутки NDC"
    .Cells(23, 1).Value = "Паллетов FG NDC"
    .Cells(24, 1).Value = "Паллетов RM NDC"
    .Cells(26, 2).Value = "Первая смена"
    .Cells(26, 27).Value = "Вторая смена"
    .Cells(26, 43).Value = "ИТОГО за сутки"
    .Cells(27, 1).Value = "Паллетов RM ИТОГО"
    .Cells(28, 1).Value = "Паллетов FG ИТОГО"
    
    Application.DisplayAlerts = False
    
    Set rng = .Range(.Cells(13, 2), .Cells(13, 8))
    rng.Merge
    rng.Interior.color = PalletsSumRGB
    Set rng = .Range(.Cells(13, 27), .Cells(13, 33))
    rng.Merge
    rng.Interior.color = PalletsSumRGB
    Set rng = .Range(.Cells(13, 43), .Cells(13, 49))
    rng.Merge
    rng.Interior.color = PalletsSumRGB
    
    Set rng = .Range(.Cells(14, 2), .Cells(14, 8))
    rng.Merge
    rng.Interior.color = PalletsSumRGB
    Set rng = .Range(.Cells(14, 27), .Cells(14, 33))
    rng.Merge
    rng.Interior.color = PalletsSumRGB
    Set rng = .Range(.Cells(14, 43), .Cells(14, 49))
    rng.Merge
    rng.Interior.color = PalletsSumRGB
    
    Set rng = .Range(.Cells(15, 2), .Cells(15, 8))
    rng.Merge
    rng.Interior.color = PalletsSumRGB
    Set rng = .Range(.Cells(15, 27), .Cells(15, 33))
    rng.Merge
    rng.Interior.color = PalletsSumRGB
    Set rng = .Range(.Cells(15, 43), .Cells(15, 49))
    rng.Merge
    rng.Interior.color = PalletsSumRGB
    
    Set rng = .Range(.Cells(22, 2), .Cells(22, 8))
    rng.Merge
    rng.Interior.color = PalletsSumRGB
    Set rng = .Range(.Cells(22, 27), .Cells(22, 33))
    rng.Merge
    rng.Interior.color = PalletsSumRGB
    Set rng = .Range(.Cells(22, 43), .Cells(22, 49))
    rng.Merge
    rng.Interior.color = PalletsSumRGB
    
    
    Set rng = .Range(.Cells(23, 2), .Cells(23, 8))
    rng.Merge
    rng.Interior.color = PalletsSumRGB
    Set rng = .Range(.Cells(23, 27), .Cells(23, 33))
    rng.Merge
    rng.Interior.color = PalletsSumRGB
    Set rng = .Range(.Cells(23, 43), .Cells(23, 49))
    rng.Merge
    rng.Interior.color = PalletsSumRGB
    
    Set rng = .Range(.Cells(24, 2), .Cells(24, 8))
    rng.Merge
    rng.Interior.color = PalletsSumRGB
    Set rng = .Range(.Cells(24, 27), .Cells(24, 33))
    rng.Merge
    rng.Interior.color = PalletsSumRGB
    Set rng = .Range(.Cells(24, 43), .Cells(24, 49))
    rng.Merge
    rng.Interior.color = PalletsSumRGB
    
    Set rng = .Range(.Cells(26, 2), .Cells(26, 8))
    rng.Merge
    rng.Interior.color = PalletsSumRGB
    Set rng = .Range(.Cells(26, 27), .Cells(26, 33))
    rng.Merge
    rng.Interior.color = PalletsSumRGB
    Set rng = .Range(.Cells(26, 43), .Cells(26, 49))
    rng.Merge
    rng.Interior.color = PalletsSumRGB
    
    Set rng = .Range(.Cells(27, 2), .Cells(27, 8))
    rng.Merge
    rng.Interior.color = PalletsSumRGB
    Set rng = .Range(.Cells(27, 27), .Cells(27, 33))
    rng.Merge
    rng.Interior.color = PalletsSumRGB
    Set rng = .Range(.Cells(27, 43), .Cells(27, 49))
    rng.Merge
    rng.Interior.color = PalletsSumRGB
    
    Set rng = .Range(.Cells(28, 2), .Cells(28, 8))
    rng.Merge
    rng.Interior.color = PalletsSumRGB
    Set rng = .Range(.Cells(28, 27), .Cells(28, 33))
    rng.Merge
    rng.Interior.color = PalletsSumRGB
    Set rng = .Range(.Cells(28, 43), .Cells(28, 49))
    rng.Merge
    rng.Interior.color = PalletsSumRGB
    
    Application.DisplayAlerts = True
    End With
    
    ws.Cells(1, 1).EntireColumn.AutoFit
    
    Call SetRangeBorder(ws.UsedRange)
End Sub
Function FLineMatcher(lineName As String) As Integer
    Select Case lineName
        Case "Линия 1"
            FLineMatcher = 2
        Case "Линия 5"
            FLineMatcher = 3
        Case "Линия 7"
            FLineMatcher = 4
        Case "Линия 8"
            FLineMatcher = 5
        Case "Линия 10"
            FLineMatcher = 6
        Case "Линия 11"
            FLineMatcher = 7
        Case "Линия 12"
            FLineMatcher = 8
        Case "Линия 13"
            FLineMatcher = 9
        Case "Линия 15"
            FLineMatcher = 10
        Case "Линия 16"
            FLineMatcher = 11
        Case "Линия 52"
            FLineMatcher = 12
        Case "Линия 3"
            FLineMatcher = 17
        Case "Линия 6"
            FLineMatcher = 18
        Case "Линия 14"
            FLineMatcher = 19
        Case "Линия 17"
            FLineMatcher = 20
        Case "Линия 56"
            FLineMatcher = 21
    End Select
End Function
Sub CreateTimeTableRecord(rng As Range, str As String, color As Long)

    Application.DisplayAlerts = False
    
    rng.Merge
    rng.Cells(1, 1) = str
    rng.Cells(1, 1).Interior.color = color
    
    Application.DisplayAlerts = True

End Sub

Private Sub SetRangeBorder(poRng As Range)
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
Function FDppType(lineName As String) As String
    If (lineName = "Линия 1") Or (lineName = "Линия 5") Or (lineName = "Линия 7") Or (lineName = "Линия 8") Or (lineName = "Линия 10") Or (lineName = "Линия 11") Or (lineName = "Линия 12") Or (lineName = "Линия 13") Or (lineName = "Линия 15") Or (lineName = "Линия 16") Or (lineName = "Линия 52") Then
        FDppType = "BAP"
    Else
        FDppType = "NDC"
    End If
End Function


