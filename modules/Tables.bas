Attribute VB_Name = "Tables"
Sub MergeTables()
Debug.Print Now
Application.ScreenUpdating = False

Dim wsTimeTableSheet As Worksheet
Dim sheetsCounter As Integer
Dim wb As Workbook: Set wb = ActiveWorkbook

For sheetsCounter = 1 To wb.Sheets.Count
    If IsDate(wb.Sheets(sheetsCounter).Name) Then
        Set wsTimeTableSheet = wb.Sheets(sheetsCounter)
        'check
        'Debug.Print (wsTimeTableSheet.Name)
        
    
    
    End If


Next sheetsCounter

Application.ScreenUpdating = True
Debug.Print Now
End Sub
Sub CreateIDTable()

Debug.Print Now
Application.ScreenUpdating = False

Dim WSRecords As Worksheet: Set WSRecords = ActiveWorkbook.Worksheets("Records")
Dim wsDate As Worksheet
Dim RecordRowCounter As Integer
Dim LineRow As Integer
Dim prevLineRow As Integer
Dim TimeTracker As Integer
Dim TimeTrackerLim As Integer
Dim DateName As String
Dim RecordLength As Integer
Dim RecordRange As Range
Dim RecordRangeTwo As Range
Dim i As Integer
Dim CleanRGB As Long: CleanRGB = RGB(255, 119, 0)
Dim NoPalletRGB As Long: NoPalletRGB = RGB(255, 0, 0)
Dim ShutdownRGB As Long: ShutdownRGB = RGB(128, 128, 128)
Dim ProductRGB As Long: ProductRGB = RGB(150, 255, 107)
Dim ChangeRGB As Long: ChangeRGB = RGB(0, 0, 255)
Dim peresmenka As Integer: peresmenka = 26
Dim potracheno As Integer
Dim prevDateName As String
Dim amountPart As Long
Dim palletPart As Long
Dim countPart As Long
Dim dppType As String
Dim vCell As Range

Dim CalculationUtils As New CalculationUtils
Dim TablesUtils As New TablesUtils

Call CalculationUtils.SDeleteTimeTableSheets

With WSRecords
TimeTracker = 1
TimeTrackeLim = 0
For RecordRowCounter = 2 To Utils.FLastRow(WSRecords)
    DateName = .Cells(RecordRowCounter, 1)
    If Not FSheetExists(DateName) Then
        Set wsDate = Utils.SCreateWS(DateName)
        Call TablesUtils.TablesFormat(wsDate)
    Else
        Set wsDate = ActiveWorkbook.Worksheets(DateName)
    End If
    
    LineRow = TablesUtils.LineMatcher(.Cells(RecordRowCounter, 2))
    dppType = CalculationUtils.FDppType(.Cells(RecordRowCounter, 2))
    
    If .Cells(RecordRowCounter, 1).Value <> .Cells(RecordRowCounter - 1, 1).Value Or .Cells(RecordRowCounter, 2) <> .Cells(RecordRowCounter - 1, 2) Then
        If TimeTrackerLim > TimeTracker - 1 Then
            Application.DisplayAlerts = False
                Set RecordRange = ActiveWorkbook.Worksheets(prevDateName).Range(ActiveWorkbook.Worksheets(prevDateName).Cells(prevLineRow, TimeTracker - 1), ActiveWorkbook.Worksheets(prevDateName).Cells(prevLineRow, TimeTrackerLim))
                'RecordRange.Merge
            Application.DisplayAlerts = True
        End If
        TimeTracker = 2
        
        Select Case .Cells(RecordRowCounter, 3)
        
            Case "�����"
                wsDate.Cells(LineRow, TimeTracker) = "�"
                'WSDate.Cells(LineRow, TimeTracker).Interior.color = CleanRGB
                wsDate.Cells(LineRow, 26) = "�"
                'WSDate.Cells(LineRow, 26).Interior.color = CleanRGB
                TimeTracker = TimeTracker + 1
                TimeTrackerLim = 49
            
            Case "�������� �����"
                Set RecordRange = wsDate.Range(wsDate.Cells(LineRow, 2), wsDate.Cells(LineRow, 25))
                Call TablesUtils.CreateSimpleRecord(RecordRange, "Shutdown")
                'For Each vCell In RecordRange
                '    vCell.Value = "Shutdown"
                'Next vCell
                'Call CalculationUtils.CreateTimeTableRecord(RecordRange, "Shutdown", ShutdownRGB)
                TimeTracker = 26
                wsDate.Cells(LineRow, TimeTracker) = "�"
                'WSDate.Cells(LineRow, TimeTracker).Interior.color = CleanRGB
                TimeTracker = TimeTracker + 1
                TimeTrackerLim = 49
            
            Case "�������� �����"
                wsDate.Cells(LineRow, TimeTracker) = "�"
                'WSDate.Cells(LineRow, TimeTracker).Interior.color = CleanRGB
                TimeTracker = TimeTracker + 1
                Set RecordRange = wsDate.Range(wsDate.Cells(LineRow, 26), wsDate.Cells(LineRow, 49))
                Call TablesUtils.CreateSimpleRecord(RecordRange, "Shutdown")
                'For Each vCell In RecordRange
                '    vCell.Value = "Shutdown"
                'Next vCell
                'Call CalculationUtils.CreateTimeTableRecord(RecordRange, "Shutdown", ShutdownRGB)
                TimeTrackerLim = 25
       End Select
    ElseIf (.Cells(RecordRowCounter, 4).Value <> .Cells(RecordRowCounter - 1, 4).Value) Then
        If (TimeTracker = peresmenka) Then
            TimeTracker = TimeTracker + 1
        End If
        wsDate.Cells(LineRow, TimeTracker) = "��"
        'WSDate.Cells(LineRow, TimeTracker).Interior.color = CleanRGB
        TimeTracker = TimeTracker + 1
        If (TimeTracker = peresmenka) Then
            TimeTracker = TimeTracker + 1
        End If
    End If
 
    RecordLength = WorksheetFunction.Round(.Cells(RecordRowCounter, 5) / .Cells(RecordRowCounter, 7) * 2, 0) + 1
    If (RecordLength = 0) Then
        RecordLength = 1
    End If
    If (RecordLength > peresmenka - TimeTracker And peresmenka > TimeTracker) Then
        'amountPart = WorksheetFunction.RoundUp(.Cells(RecordRowCounter, 5) * (1 - (peresmenka - TimeTracker) / RecordLength), 0)
        'palletPart = WorksheetFunction.RoundUp(.Cells(RecordRowCounter, 6) * (1 - (peresmenka - TimeTracker) / RecordLength), 0)
        'countPart = WorksheetFunction.RoundUp(.Cells(RecordRowCounter, 8) * (1 - (peresmenka - TimeTracker) / RecordLength), 0)
        'If dppType = "BAP" Then
        '    WSDate.Cells(14, 2).Value = WSDate.Cells(14, 2).Value + .Cells(RecordRowCounter, 8) - countPart
        '    WSDate.Cells(14, peresmenka + 1).Value = WSDate.Cells(14, peresmenka + 1).Value + countPart
        '
        '    WSDate.Cells(15, 2).Value = WSDate.Cells(15, 2).Value + .Cells(RecordRowCounter, 6) - palletPart
        '    WSDate.Cells(15, peresmenka + 1).Value = WSDate.Cells(15, peresmenka + 1).Value + palletPart
        'Else
        '    WSDate.Cells(23, 2).Value = WSDate.Cells(23, 2).Value + .Cells(RecordRowCounter, 8) - countPart
        '    WSDate.Cells(23, peresmenka + 1).Value = WSDate.Cells(23, peresmenka + 1).Value + countPart
        '
        '    WSDate.Cells(24, 2).Value = WSDate.Cells(24, 2).Value + .Cells(RecordRowCounter, 6) - palletPart
        '    WSDate.Cells(24, peresmenka + 1).Value = WSDate.Cells(24, peresmenka + 1).Value + palletPart
        'End If
        Set RecordRange = wsDate.Range(wsDate.Cells(LineRow, TimeTracker), wsDate.Cells(LineRow, peresmenka - 1))
        If (RecordLength + TimeTracker > TimeTrackerLim) Then
            Set RecordRangeTwo = wsDate.Range(wsDate.Cells(LineRow, TimeTracker), wsDate.Cells(LineRow, TimeTrackerLim))
        Else
            RecordLength = RecordLength - (peresmenka - TimeTracker)
            Set RecordRangeTwo = wsDate.Range(wsDate.Cells(LineRow, peresmenka + 1), wsDate.Cells(LineRow, peresmenka + RecordLength))
        End If
        If (.Cells(RecordRowCounter, 6) > 0) Then
            Call TablesUtils.CreateIDRecord(RecordRange, .Cells(RecordRowCounter, 4))
            Call TablesUtils.CreateIDRecord(RecordRangeTwo, .Cells(RecordRowCounter, 4))
        '    Call CalculationUtils.CreateTimeTableRecord(RecordRange, "Id:" & .Cells(RecordRowCounter, 4) & "  FG,kg:" & (.Cells(RecordRowCounter, 5) - amountPart) & "  FG,pal:" & (.Cells(RecordRowCounter, 8) - countPart) & " RM,pal:" & (.Cells(RecordRowCounter, 6) - palletPart), ProductRGB)
        '    Call CalculationUtils.CreateTimeTableRecord(RecordRangeTwo, "Id:" & .Cells(RecordRowCounter, 4) & "  FG,kg:" & amountPart & "  FG,pal:" & countPart & " RM,pal:" & palletPart, ProductRGB)
        Else
            Call TablesUtils.CreateIDRecord(RecordRange, .Cells(RecordRowCounter, 4))
            Call TablesUtils.CreateIDRecord(RecordRangeTwo, .Cells(RecordRowCounter, 4))
        '    Call CalculationUtils.CreateTimeTableRecord(RecordRange, "Id:" & .Cells(RecordRowCounter, 4) & "  FG,kg:" & (.Cells(RecordRowCounter, 5) - amountPart) & "  FG,pal:" & (.Cells(RecordRowCounter, 8) - countPart) & " RM,pal:" & (.Cells(RecordRowCounter, 6) - palletPart), NoPalletRGB)
        '    Call CalculationUtils.CreateTimeTableRecord(RecordRangeTwo, "Id:" & .Cells(RecordRowCounter, 4) & "  FG,kg:" & amountPart & " FG, pal:" & countPart & " RM,pal:" & palletPart, NoPalletRGB)
        End If
        TimeTracker = TimeTracker + WorksheetFunction.Round(.Cells(RecordRowCounter, 5) / .Cells(RecordRowCounter, 7) * 2, 0) + 1
    Else
    '    If (peresmenka > TimeTracker) Then
    '        If dppType = "BAP" Then
    '            WSDate.Cells(14, 2) = WSDate.Cells(14, 2).Value + .Cells(RecordRowCounter, 8)
    '            WSDate.Cells(15, 2) = WSDate.Cells(15, 2).Value + .Cells(RecordRowCounter, 6)
    '        Else
    '            WSDate.Cells(23, 2) = WSDate.Cells(23, 2).Value + .Cells(RecordRowCounter, 8)
    '            WSDate.Cells(24, 2) = WSDate.Cells(24, 2).Value + .Cells(RecordRowCounter, 6)
    '        End If
    '    Else
    '        If dppType = "BAP" Then
    '            WSDate.Cells(14, peresmenka + 1) = WSDate.Cells(14, peresmenka + 1).Value + .Cells(RecordRowCounter, 8)
    '            WSDate.Cells(15, peresmenka + 1) = WSDate.Cells(15, peresmenka + 1).Value + .Cells(RecordRowCounter, 6)
    '        Else
    '            WSDate.Cells(23, peresmenka + 1) = WSDate.Cells(23, peresmenka + 1).Value + .Cells(RecordRowCounter, 8)
    '            WSDate.Cells(24, peresmenka + 1) = WSDate.Cells(24, peresmenka + 1).Value + .Cells(RecordRowCounter, 6)
    '         End If
    '    End If
        If (RecordLength + TimeTracker > TimeTrackerLim) Then
            Set RecordRange = wsDate.Range(wsDate.Cells(LineRow, TimeTracker), wsDate.Cells(LineRow, TimeTrackerLim))
        Else
            Set RecordRange = wsDate.Range(wsDate.Cells(LineRow, TimeTracker), wsDate.Cells(LineRow, TimeTracker + RecordLength - 1))
        End If
        If (.Cells(RecordRowCounter, 6) > 0) Then
            Call TablesUtils.CreateIDRecord(RecordRange, .Cells(RecordRowCounter, 4))
    '        Call CalculationUtils.CreateTimeTableRecord(RecordRange, "Id:" & .Cells(RecordRowCounter, 4) & "  FG,kg:" & .Cells(RecordRowCounter, 5) & " FG,pal:" & .Cells(RecordRowCounter, 8) & " RM,pal:" & .Cells(RecordRowCounter, 6), ProductRGB)
        Else
            Call TablesUtils.CreateIDRecord(RecordRange, .Cells(RecordRowCounter, 4))
    '        Call CalculationUtils.CreateTimeTableRecord(RecordRange, "Id:" & .Cells(RecordRowCounter, 4) & "  FG,kg:" & .Cells(RecordRowCounter, 5) & " FG,pal:" & .Cells(RecordRowCounter, 8) & " RM,pal:" & .Cells(RecordRowCounter, 6), NoPalletRGB)
        End If
        TimeTracker = TimeTracker + RecordLength
    End If
    
    prevDateName = DateName
    prevLineRow = LineRow
    
    'WSDate.Cells(14, 43).Value = WSDate.Cells(14, 2).Value + WSDate.Cells(14, peresmenka + 1)
    'WSDate.Cells(15, 43).Value = WSDate.Cells(15, 2).Value + WSDate.Cells(15, peresmenka + 1)
    'WSDate.Cells(23, 43).Value = WSDate.Cells(23, 2).Value + WSDate.Cells(23, peresmenka + 1)
    'WSDate.Cells(24, 43).Value = WSDate.Cells(24, 2).Value + WSDate.Cells(24, peresmenka + 1)
    
    'WSDate.Cells(27, 2).Value = WSDate.Cells(14, 2).Value + WSDate.Cells(23, 2)
    'WSDate.Cells(28, 2).Value = WSDate.Cells(15, 2).Value + WSDate.Cells(24, 2)
    'WSDate.Cells(27, peresmenka + 1).Value = WSDate.Cells(14, peresmenka + 1).Value + WSDate.Cells(23, peresmenka + 1)
    'WSDate.Cells(28, peresmenka + 1).Value = WSDate.Cells(15, peresmenka + 1).Value + WSDate.Cells(24, peresmenka + 1)
    'WSDate.Cells(27, peresmenka + 1).Value = WSDate.Cells(14, peresmenka + 1).Value + WSDate.Cells(23, peresmenka + 1)
    'WSDate.Cells(28, peresmenka + 1).Value = WSDate.Cells(15, peresmenka + 1).Value + WSDate.Cells(24, peresmenka + 1)
    'WSDate.Cells(27, 43).Value = WSDate.Cells(14, 43).Value + WSDate.Cells(23, 43)
    'WSDate.Cells(28, 43).Value = WSDate.Cells(15, 43).Value + WSDate.Cells(24, 43)
    

Next RecordRowCounter
If TimeTrackerLim > TimeTracker - 1 Then
    Application.DisplayAlerts = False
        Set RecordRange = ActiveWorkbook.Worksheets(prevDateName).Range(ActiveWorkbook.Worksheets(prevDateName).Cells(prevLineRow, TimeTracker - 1), ActiveWorkbook.Worksheets(prevDateName).Cells(prevLineRow, TimeTrackerLim))
        'RecordRange.Merge
    Application.DisplayAlerts = True
End If
        
End With

Debug.Print Now
Application.ScreenUpdating = True
End Sub

Sub CreateFGTable()

Debug.Print Now
Application.ScreenUpdating = False

Dim WSRecords As Worksheet: Set WSRecords = ActiveWorkbook.Worksheets("Records")
Dim wsDate As Worksheet
Dim RecordRowCounter As Integer
Dim LineRow As Integer
Dim prevLineRow As Integer
Dim TimeTracker As Integer
Dim TimeTrackerLim As Integer
Dim DateName As String
Dim RecordLength As Integer
Dim RecordRange As Range
Dim RecordRangeTwo As Range
Dim i As Integer
Dim CleanRGB As Long: CleanRGB = RGB(255, 119, 0)
Dim NoPalletRGB As Long: NoPalletRGB = RGB(255, 0, 0)
Dim ShutdownRGB As Long: ShutdownRGB = RGB(128, 128, 128)
Dim ProductRGB As Long: ProductRGB = RGB(150, 255, 107)
Dim ChangeRGB As Long: ChangeRGB = RGB(0, 0, 255)
Dim peresmenka As Integer: peresmenka = 26
Dim potracheno As Integer
Dim prevDateName As String
Dim amountPart As Long
Dim palletPart As Long
Dim countPart As Long
Dim dppType As String
Dim vCell As Range

Dim CalculationUtils As New CalculationUtils
Dim TablesUtils As New TablesUtils

Call CalculationUtils.SDeleteTimeTableSheets

With WSRecords
TimeTracker = 1
TimeTrackeLim = 0
For RecordRowCounter = 2 To Utils.FLastRow(WSRecords)
    DateName = .Cells(RecordRowCounter, 1)
    If Not FSheetExists(DateName) Then
        Set wsDate = Utils.SCreateWS(DateName)
        Call TablesUtils.TablesFormat(wsDate)
    Else
        Set wsDate = ActiveWorkbook.Worksheets(DateName)
    End If
    
    LineRow = TablesUtils.LineMatcher(.Cells(RecordRowCounter, 2))
    dppType = CalculationUtils.FDppType(.Cells(RecordRowCounter, 2))
    
    If .Cells(RecordRowCounter, 1).Value <> .Cells(RecordRowCounter - 1, 1).Value Or .Cells(RecordRowCounter, 2) <> .Cells(RecordRowCounter - 1, 2) Then
        If TimeTrackerLim > TimeTracker - 1 Then
            Application.DisplayAlerts = False
                Set RecordRange = ActiveWorkbook.Worksheets(prevDateName).Range(ActiveWorkbook.Worksheets(prevDateName).Cells(prevLineRow, TimeTracker - 1), ActiveWorkbook.Worksheets(prevDateName).Cells(prevLineRow, TimeTrackerLim))
                'RecordRange.Merge
            Application.DisplayAlerts = True
        End If
        TimeTracker = 2
        
        Select Case .Cells(RecordRowCounter, 3)
        
            Case "�����"
                wsDate.Cells(LineRow, TimeTracker) = "�"
                'WSDate.Cells(LineRow, TimeTracker).Interior.color = CleanRGB
                wsDate.Cells(LineRow, 26) = "�"
                'WSDate.Cells(LineRow, 26).Interior.color = CleanRGB
                TimeTracker = TimeTracker + 1
                TimeTrackerLim = 49
            
            Case "�������� �����"
                Set RecordRange = wsDate.Range(wsDate.Cells(LineRow, 2), wsDate.Cells(LineRow, 25))
                Call TablesUtils.CreateSimpleRecord(RecordRange, "Shutdown")
                'For Each vCell In RecordRange
                '    vCell.Value = "Shutdown"
                'Next vCell
                'Call CalculationUtils.CreateTimeTableRecord(RecordRange, "Shutdown", ShutdownRGB)
                TimeTracker = 26
                wsDate.Cells(LineRow, TimeTracker) = "�"
                'WSDate.Cells(LineRow, TimeTracker).Interior.color = CleanRGB
                TimeTracker = TimeTracker + 1
                TimeTrackerLim = 49
            
            Case "�������� �����"
                wsDate.Cells(LineRow, TimeTracker) = "�"
                'WSDate.Cells(LineRow, TimeTracker).Interior.color = CleanRGB
                TimeTracker = TimeTracker + 1
                Set RecordRange = wsDate.Range(wsDate.Cells(LineRow, 26), wsDate.Cells(LineRow, 49))
                Call TablesUtils.CreateSimpleRecord(RecordRange, "Shutdown")
                'For Each vCell In RecordRange
                '    vCell.Value = "Shutdown"
                'Next vCell
                'Call CalculationUtils.CreateTimeTableRecord(RecordRange, "Shutdown", ShutdownRGB)
                TimeTrackerLim = 25
       End Select
    ElseIf (.Cells(RecordRowCounter, 4).Value <> .Cells(RecordRowCounter - 1, 4).Value) Then
        If (TimeTracker = peresmenka) Then
            TimeTracker = TimeTracker + 1
        End If
        wsDate.Cells(LineRow, TimeTracker) = "��"
        'WSDate.Cells(LineRow, TimeTracker).Interior.color = CleanRGB
        TimeTracker = TimeTracker + 1
        If (TimeTracker = peresmenka) Then
            TimeTracker = TimeTracker + 1
        End If
    End If
 
    RecordLength = WorksheetFunction.Round(.Cells(RecordRowCounter, 5).Value * 2 / .Cells(RecordRowCounter, 7).Value, 0) + 1
    If (RecordLength = 0) Then
        RecordLength = 1
    End If
    If (RecordLength > peresmenka - TimeTracker And peresmenka > TimeTracker) Then
        'amountPart = WorksheetFunction.RoundUp(.Cells(RecordRowCounter, 5) * (1 - (peresmenka - TimeTracker) / RecordLength), 0)
        'palletPart = WorksheetFunction.RoundUp(.Cells(RecordRowCounter, 6) * (1 - (peresmenka - TimeTracker) / RecordLength), 0)
        'countPart = WorksheetFunction.RoundUp(.Cells(RecordRowCounter, 8) * (1 - (peresmenka - TimeTracker) / RecordLength), 0)
        'If dppType = "BAP" Then
        '    WSDate.Cells(14, 2).Value = WSDate.Cells(14, 2).Value + .Cells(RecordRowCounter, 8) - countPart
        '    WSDate.Cells(14, peresmenka + 1).Value = WSDate.Cells(14, peresmenka + 1).Value + countPart
        '
        '    WSDate.Cells(15, 2).Value = WSDate.Cells(15, 2).Value + .Cells(RecordRowCounter, 6) - palletPart
        '    WSDate.Cells(15, peresmenka + 1).Value = WSDate.Cells(15, peresmenka + 1).Value + palletPart
        'Else
        '    WSDate.Cells(23, 2).Value = WSDate.Cells(23, 2).Value + .Cells(RecordRowCounter, 8) - countPart
        '    WSDate.Cells(23, peresmenka + 1).Value = WSDate.Cells(23, peresmenka + 1).Value + countPart
        '
        '    WSDate.Cells(24, 2).Value = WSDate.Cells(24, 2).Value + .Cells(RecordRowCounter, 6) - palletPart
        '    WSDate.Cells(24, peresmenka + 1).Value = WSDate.Cells(24, peresmenka + 1).Value + palletPart
        'End If
        Set RecordRange = wsDate.Range(wsDate.Cells(LineRow, TimeTracker), wsDate.Cells(LineRow, peresmenka - 1))
        If (RecordLength + TimeTracker > TimeTrackerLim) Then
            Set RecordRangeTwo = wsDate.Range(wsDate.Cells(LineRow, TimeTracker), wsDate.Cells(LineRow, TimeTrackerLim))
        Else
            RecordLength = RecordLength - (peresmenka - TimeTracker)
            Set RecordRangeTwo = wsDate.Range(wsDate.Cells(LineRow, peresmenka + 1), wsDate.Cells(LineRow, peresmenka + RecordLength))
        End If
        If (.Cells(RecordRowCounter, 6) > 0) Then
            Call TablesUtils.CreateFGRecord(RecordRange, .Cells(RecordRowCounter, 5), .Cells(RecordRowCounter, 7))
            Call TablesUtils.CreateFGRecord(RecordRangeTwo, .Cells(RecordRowCounter, 5), .Cells(RecordRowCounter, 7))
            
            'Call TablesUtils.CreateIDRecord(RecordRange, .Cells(RecordRowCounter, 4))
            'Call TablesUtils.CreateIDRecord(RecordRangeTwo, .Cells(RecordRowCounter, 4))
        
        
        
        '    Call CalculationUtils.CreateTimeTableRecord(RecordRange, "Id:" & .Cells(RecordRowCounter, 4) & "  FG,kg:" & (.Cells(RecordRowCounter, 5) - amountPart) & "  FG,pal:" & (.Cells(RecordRowCounter, 8) - countPart) & " RM,pal:" & (.Cells(RecordRowCounter, 6) - palletPart), ProductRGB)
        '    Call CalculationUtils.CreateTimeTableRecord(RecordRangeTwo, "Id:" & .Cells(RecordRowCounter, 4) & "  FG,kg:" & amountPart & "  FG,pal:" & countPart & " RM,pal:" & palletPart, ProductRGB)
        Else
            Call TablesUtils.CreateFGRecord(RecordRange, .Cells(RecordRowCounter, 5), .Cells(RecordRowCounter, 7))
            Call TablesUtils.CreateFGRecord(RecordRangeTwo, .Cells(RecordRowCounter, 5), .Cells(RecordRowCounter, 7))
            
            'Call TablesUtils.CreateIDRecord(RecordRange, .Cells(RecordRowCounter, 4))
            'Call TablesUtils.CreateIDRecord(RecordRangeTwo, .Cells(RecordRowCounter, 4))
            
            
        '    Call CalculationUtils.CreateTimeTableRecord(RecordRange, "Id:" & .Cells(RecordRowCounter, 4) & "  FG,kg:" & (.Cells(RecordRowCounter, 5) - amountPart) & "  FG,pal:" & (.Cells(RecordRowCounter, 8) - countPart) & " RM,pal:" & (.Cells(RecordRowCounter, 6) - palletPart), NoPalletRGB)
        '    Call CalculationUtils.CreateTimeTableRecord(RecordRangeTwo, "Id:" & .Cells(RecordRowCounter, 4) & "  FG,kg:" & amountPart & " FG, pal:" & countPart & " RM,pal:" & palletPart, NoPalletRGB)
        End If
        TimeTracker = TimeTracker + WorksheetFunction.Round(.Cells(RecordRowCounter, 5) / .Cells(RecordRowCounter, 7) * 2, 0) + 1
    Else
    '    If (peresmenka > TimeTracker) Then
    '        If dppType = "BAP" Then
    '            WSDate.Cells(14, 2) = WSDate.Cells(14, 2).Value + .Cells(RecordRowCounter, 8)
    '            WSDate.Cells(15, 2) = WSDate.Cells(15, 2).Value + .Cells(RecordRowCounter, 6)
    '        Else
    '            WSDate.Cells(23, 2) = WSDate.Cells(23, 2).Value + .Cells(RecordRowCounter, 8)
    '            WSDate.Cells(24, 2) = WSDate.Cells(24, 2).Value + .Cells(RecordRowCounter, 6)
    '        End If
    '    Else
    '        If dppType = "BAP" Then
    '            WSDate.Cells(14, peresmenka + 1) = WSDate.Cells(14, peresmenka + 1).Value + .Cells(RecordRowCounter, 8)
    '            WSDate.Cells(15, peresmenka + 1) = WSDate.Cells(15, peresmenka + 1).Value + .Cells(RecordRowCounter, 6)
    '        Else
    '            WSDate.Cells(23, peresmenka + 1) = WSDate.Cells(23, peresmenka + 1).Value + .Cells(RecordRowCounter, 8)
    '            WSDate.Cells(24, peresmenka + 1) = WSDate.Cells(24, peresmenka + 1).Value + .Cells(RecordRowCounter, 6)
    '         End If
    '    End If
        If (RecordLength + TimeTracker > TimeTrackerLim) Then
            Set RecordRange = wsDate.Range(wsDate.Cells(LineRow, TimeTracker), wsDate.Cells(LineRow, TimeTrackerLim))
        Else
            Set RecordRange = wsDate.Range(wsDate.Cells(LineRow, TimeTracker), wsDate.Cells(LineRow, TimeTracker + RecordLength - 1))
        End If
        If (.Cells(RecordRowCounter, 6) > 0) Then
            Call TablesUtils.CreateFGRecord(RecordRange, .Cells(RecordRowCounter, 5), .Cells(RecordRowCounter, 7))
            
    '        Call TablesUtils.CreateIDRecord(RecordRange, .Cells(RecordRowCounter, 4))
    '        Call CalculationUtils.CreateTimeTableRecord(RecordRange, "Id:" & .Cells(RecordRowCounter, 4) & "  FG,kg:" & .Cells(RecordRowCounter, 5) & " FG,pal:" & .Cells(RecordRowCounter, 8) & " RM,pal:" & .Cells(RecordRowCounter, 6), ProductRGB)
        Else
            Call TablesUtils.CreateFGRecord(RecordRange, .Cells(RecordRowCounter, 5), .Cells(RecordRowCounter, 7))
            
            
    '        Call TablesUtils.CreateIDRecord(RecordRange, .Cells(RecordRowCounter, 4))
    '        Call CalculationUtils.CreateTimeTableRecord(RecordRange, "Id:" & .Cells(RecordRowCounter, 4) & "  FG,kg:" & .Cells(RecordRowCounter, 5) & " FG,pal:" & .Cells(RecordRowCounter, 8) & " RM,pal:" & .Cells(RecordRowCounter, 6), NoPalletRGB)
        End If
        TimeTracker = TimeTracker + RecordLength
    End If
    
    prevDateName = DateName
    prevLineRow = LineRow
    
    'WSDate.Cells(14, 43).Value = WSDate.Cells(14, 2).Value + WSDate.Cells(14, peresmenka + 1)
    'WSDate.Cells(15, 43).Value = WSDate.Cells(15, 2).Value + WSDate.Cells(15, peresmenka + 1)
    'WSDate.Cells(23, 43).Value = WSDate.Cells(23, 2).Value + WSDate.Cells(23, peresmenka + 1)
    'WSDate.Cells(24, 43).Value = WSDate.Cells(24, 2).Value + WSDate.Cells(24, peresmenka + 1)
    
    'WSDate.Cells(27, 2).Value = WSDate.Cells(14, 2).Value + WSDate.Cells(23, 2)
    'WSDate.Cells(28, 2).Value = WSDate.Cells(15, 2).Value + WSDate.Cells(24, 2)
    'WSDate.Cells(27, peresmenka + 1).Value = WSDate.Cells(14, peresmenka + 1).Value + WSDate.Cells(23, peresmenka + 1)
    'WSDate.Cells(28, peresmenka + 1).Value = WSDate.Cells(15, peresmenka + 1).Value + WSDate.Cells(24, peresmenka + 1)
    'WSDate.Cells(27, peresmenka + 1).Value = WSDate.Cells(14, peresmenka + 1).Value + WSDate.Cells(23, peresmenka + 1)
    'WSDate.Cells(28, peresmenka + 1).Value = WSDate.Cells(15, peresmenka + 1).Value + WSDate.Cells(24, peresmenka + 1)
    'WSDate.Cells(27, 43).Value = WSDate.Cells(14, 43).Value + WSDate.Cells(23, 43)
    'WSDate.Cells(28, 43).Value = WSDate.Cells(15, 43).Value + WSDate.Cells(24, 43)
    

Next RecordRowCounter
If TimeTrackerLim > TimeTracker - 1 Then
    Application.DisplayAlerts = False
        Set RecordRange = ActiveWorkbook.Worksheets(prevDateName).Range(ActiveWorkbook.Worksheets(prevDateName).Cells(prevLineRow, TimeTracker - 1), ActiveWorkbook.Worksheets(prevDateName).Cells(prevLineRow, TimeTrackerLim))
        'RecordRange.Merge
    Application.DisplayAlerts = True
End If
        
End With

Debug.Print Now
Application.ScreenUpdating = True
End Sub

Sub testMerge()
Dim TablesUtils As New TablesUtils
Call TablesUtils.MergeTables("FG")
End Sub



