Attribute VB_Name = "Utils"
Sub SDeleteOrdersSheets(Optional dpp As String)
    Dim i As Integer
    Application.DisplayAlerts = False
    For i = Sheets.Count To 1 Step -1
        If (InStr(Sheets(i).Name, "ордера") <> 0) And InStr(Sheets(i).Name, dpp) <> 0 Then
            Sheets(i).Delete
        End If
    Next i
    Application.DisplayAlerts = True
End Sub
Function getOrderSheetName(weekNumber, Optional wb As Workbook, Optional wsDpp As Worksheet) As String
    Dim cwNumber As Integer
    Dim weekRange As Range
    Dim DictionaryUtils As New DictionaryUtils
    
    If wb Is Nothing Then Set wb = ActiveWorkbook
    DictionaryUtils.weekNumber = weekNumber
    Set weekRange = DictionaryUtils.FWeekRange(wsDpp)
    cwNumber = weekRange.Cells(1, 2).Value
    getOrderSheetName = "ордера w" & cwNumber
End Function
Function FLastRow(Optional ws As Worksheet) As Integer
    If ws Is Nothing Then Set ws = ActiveWorkbook.Worksheets("DPP")
    FLastRow = ws.UsedRange.Rows.Count
End Function
Function FLastColumn(Optional ws As Worksheet) As Integer
    If ws Is Nothing Then Set ws = ActiveWorkbook.Worksheets("DPP")
    FLastColumn = ws.UsedRange.Columns.Count
End Function
Function FSheetExists(sheetName As String, Optional wb As Workbook) As Boolean
    Dim ws As Worksheet
     If wb Is Nothing Then Set wb = ActiveWorkbook
     On Error Resume Next
     Set ws = wb.Sheets(sheetName)
     On Error GoTo 0
     FSheetExists = Not ws Is Nothing
 End Function
Function IsPivotReadyToGenerate(Optional wb As Workbook) As Boolean
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If (FSheetExists("Справочник RM") And FSheetExists("Справочник расходов")) Then
        IsPivotReadyToGenerate = True
    Else
        IsPivotReadyToGenerate = False
    End If
End Function
Function CreateSheet(shtName As String) As Worksheet
    If FSheetExists(shtName) Then
        Application.DisplayAlerts = False
        Sheets(shtName).Delete
        Application.DisplayAlerts = True
    End If
    Set CreateSheet = SCreateWS(shtName)
End Function
Function SCreateWS(shtName As String) As Worksheet
    With ActiveWorkbook
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = shtName
    End With
    Set SCreateWS = ActiveWorkbook.Worksheets(shtName)
End Function
Function IsDictionaryReadyToCalculate(Optional wb As Workbook) As Boolean
    Dim dppBapExists As Boolean
    Dim dppNdcExists As Boolean
    Dim dppBapOrd1Exists As Boolean
    Dim dppBapOrd2Exists As Boolean
    Dim dppNdcOrd1Exists As Boolean
    Dim dppNdcOrd2Exists As Boolean
    
    dppBapExists = FSheetExists("DPP_BAP")
    dppNdcExists = FSheetExists("DPP_NDC")
    If dppBapExists Then
    dppBapOrd1Exists = FSheetExists(getOrderSheetName(1, wb, Sheets("DPP_BAP")) & " BAP")
    dppBapOrd2Exists = FSheetExists(getOrderSheetName(2, wb, Sheets("DPP_BAP")) & " BAP")
    End If
    If dppNdcExists Then
    dppNdcOrd1Exists = FSheetExists(getOrderSheetName(1, wb, Sheets("DPP_NDC")) & " NDC")
    dppNdcOrd2Exists = FSheetExists(getOrderSheetName(2, wb, Sheets("DPP_NDC")) & " NDC")
    End If
    
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If FSheetExists("Pivot") Then
        If (dppBapExists And (dppBapOrd1Exists Or dppBapOrd2Exists)) Or (dppNdcExists And (dppNdcOrd1Exists Or dppNdcOrd2Exists)) Then
            IsDictionaryReadyToCalculate = True
        Else
            IsDictionaryReadyToCalculate = False
        End If
    Else
        IsDictionaryReadyToCalculate = False
    End If
End Function
Function IsAtLeastOneDateInSheets(Optional wb As Workbook) As Boolean
    If wb Is Nothing Then Set wb = ActiveWorkbook
    IsAtLeastOneDateInSheets = False
    Dim i As Integer
    If (FSheetExists("Records")) Then
        For i = Sheets.Count To 1 Step -1
            If (IsDate(Sheets(i).Name)) Then
                If (DateDiff("d", Sheets(i).Name, wb.Sheets("Records").Cells(2, 1).Value) = 0) Then
                    IsAtLeastOneDateInSheets = True
                Exit For
                End If
            End If
        Next i
    Else
        IsAtLeastOneDateInSheets = False
    End If
End Function
Sub RMFormat()
Dim ws As Worksheet
Set ws = ActiveWorkbook.Sheets("Справочник RM")
ws.Cells(1, 5) = "Common material"
ws.Cells(1, 5).Font.Bold = True
Columns("E").AutoFit
End Sub
