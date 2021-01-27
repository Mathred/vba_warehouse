Attribute VB_Name = "Dictionary"
Sub GenerateRecordsTable()
Application.ScreenUpdating = False
Debug.Print Now
Dim WSOrders As Worksheet

Dim WsDppNdc As Worksheet
Dim WsDppBap As Worksheet
Dim WsDppArray(1 To 2) As Worksheet
Dim WsDppCounter As Integer
Dim WsDppCounterLimit As Integer

If FSheetExists("DPP_NDC") Then
   Set WsDppNdc = ActiveWorkbook.Sheets("DPP_NDC")
End If

If FSheetExists("DPP_BAP") Then
   Set WsDppBap = ActiveWorkbook.Sheets("DPP_BAP")
End If

If FSheetExists("DPP_BAP") And FSheetExists("DPP_NDC") Then
    Set WsDppArray(1) = WsDppBap
    Set WsDppArray(2) = WsDppNdc
    WsDppCounterLimit = 2
ElseIf FSheetExists("DPP_BAP") Then
    Set WsDppArray(1) = WsDppBap
    WsDppCounterLimit = 1
ElseIf FSheetExists("DPP_NDC") Then
    Set WsDppArray(1) = WsDppNdc
    WsDppCounterLimit = 1
End If







Dim wsDpp As Worksheet
Dim wsDate As Worksheet
Dim lineName As String
Dim dat As Date
Dim linerange As Range
Dim daterange As Range
Dim prevLineRange As Range
Dim product As Long
Dim orderRow As Integer
Dim recordRow As Integer
Dim orderlistname As String
Dim weekNumber As Integer
Dim fddp2 As Long
Dim recordsSheet As Worksheet

Dim DictionaryUtils As New DictionaryUtils

Set recordsSheet = Utils.CreateSheet("Records")
Call DictionaryUtils.RecordsHeader(recordsSheet)

recordRow = 2

For WsDppCounter = 1 To WsDppCounterLimit
    Set wsDpp = WsDppArray(WsDppCounter)
    If WsDppCounter = 2 Then
        Debug.Print "Break"
    End If
    For weekNumber = 1 To 2
    
        
        DictionaryUtils.weekNumber = weekNumber
        
        If wsDpp.Name = "DPP_BAP" Then
            orderlistname = "ордера w" & DictionaryUtils.FWeekRange(wsDpp).Cells(1, 2).Value & " BAP"
        ElseIf wsDpp.Name = "DPP_NDC" Then
            orderlistname = "ордера w" & DictionaryUtils.FWeekRange(wsDpp).Cells(1, 2).Value & " NDC"
        End If
        
        If Utils.FSheetExists(orderlistname) Then
        
            Set WSOrders = ActiveWorkbook.Sheets(orderlistname)
            
            For orderRow = 2 To Utils.FLastRow(WSOrders)
                If Not IsEmpty(WSOrders.Cells(orderRow, 1)) Then
                    lineName = DictionaryUtils.FLineRenamer(WSOrders.Cells(orderRow, 1))
                    Set linerange = DictionaryUtils.FLineRange(lineName, wsDpp, prevLineRange)
                ElseIf IsDate(WSOrders.Cells(orderRow, 2)) Then
                    dat = WSOrders.Cells(orderRow, 2)
                    Set daterange = DictionaryUtils.FDateRange(dat, wsDpp)
                ElseIf (Not IsEmpty(WSOrders.Cells(orderRow, 2)) And IsNumeric(WSOrders.Cells(orderRow, 2))) Then
                    product = WSOrders.Cells(orderRow, 3).Value
                    fddp2 = DictionaryUtils.FDPP2(product, linerange, daterange, wsDpp)
                    If (fddp2 > 0) And (wsDpp.Cells(DictionaryUtils.FProductRow(linerange, product, 1, wsDpp), DictionaryUtils.FNetCol(wsDpp).Column).Value <> 0) Then
                        recordsSheet.Cells(recordRow, 1) = dat
                        recordsSheet.Cells(recordRow, 2) = lineName
                        recordsSheet.Cells(recordRow, 3) = wsDpp.Cells(DictionaryUtils.FShiftRow(linerange, wsDpp).row, daterange.Column)
                        recordsSheet.Cells(recordRow, 4) = product
                        recordsSheet.Cells(recordRow, 5) = fddp2
                        recordsSheet.Cells(recordRow, 6) = DictionaryUtils.FPalletsCount(product, Worksheets("Records").Cells(recordRow, 5).Value)
                        recordsSheet.Cells(recordRow, 7) = wsDpp.Cells(DictionaryUtils.FProductRow(linerange, product, 1, wsDpp), DictionaryUtils.FCapacityCol(wsDpp).Column)
                        recordsSheet.Cells(recordRow, 8) = Application.WorksheetFunction.RoundUp(recordsSheet.Cells(recordRow, 5).Value / wsDpp.Cells(DictionaryUtils.FProductRow(linerange, product, 1, wsDpp), DictionaryUtils.FNetCol(wsDpp).Column).Value, 0)
                        recordRow = recordRow + 1
                    End If
             
                End If
                Set prevLineRange = linerange
            Next orderRow
        Else
            MsgBox (orderlistname & " не найден")
        End If
        Set prevLineRange = Nothing
    Next weekNumber

Next WsDppCounter

Call DictionaryUtils.RecordsAutofit(recordsSheet)
Debug.Print Now
Application.ScreenUpdating = True
End Sub

