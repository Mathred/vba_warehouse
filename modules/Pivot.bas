Attribute VB_Name = "Pivot"
Function GeneratePivot() As Boolean
    
    Dim PivotSheet As Worksheet
    Dim RMSheet As Worksheet
    Dim ConsumptionSheet As Worksheet
    Dim consumptionColumn As Integer
    Dim consumptionRow As Integer
    Dim componentId As Long
    Dim productId As Long
    Dim pcsPerPallet As Long
    Set PivotSheet = Utils.CreateSheet("Pivot")
    Set RMSheet = ActiveWorkbook.Sheets("Справочник RM")
    Set ConsumptionSheet = ActiveWorkbook.Sheets("Справочник расходов")
    
    For consumptionRow = 5 To Utils.FLastRow(ConsumptionSheet) - 1
        consumptionColumn = 4
        componentId = ConsumptionSheet.Cells(consumptionRow, 1).Value
        If componentId = RMSheet.Cells(consumptionRow - 3, 1) Or componentId = RMSheet.Cells(consumptionRow - 3 + 1, 1) Then
            If (componentId <> RMSheet.Cells(consumptionRow - 3, 1)) Then
                RMSheet.Rows(consumptionRow - 3).Delete
            End If
            PivotSheet.Cells(1, consumptionRow - 2) = componentId
            For consumptionColumn = 4 To Utils.FLastColumn(ConsumptionSheet) - 1
                productId = ConsumptionSheet.Cells(2, consumptionColumn).Value
                If (consumptionRow = 5) Then
                    PivotSheet.Cells(consumptionColumn, 1) = productId
                End If
                pcsPerPallet = RMSheet.Cells(consumptionRow - 3, 4)
                If (Not IsEmpty(RMSheet.Cells(consumptionRow - 3, 5)) Or pcsPerPallet = 0) Then
                    PivotSheet.Cells(consumptionColumn, consumptionRow - 2) = 0
                Else
                    PivotSheet.Cells(consumptionColumn, consumptionRow - 2) = ConsumptionSheet.Cells(consumptionRow, consumptionColumn) / 1000 / pcsPerPallet
                End If
            Next consumptionColumn
        Else
            MsgBox ("У вас неверная пара справочников RM и расходов. Компонент " & componentId & " не найден в справочнике RM")
            Application.DisplayAlerts = False
            PivotSheet.Delete
            Application.DisplayAlerts = True
            GeneratePivot = False
            Exit Function
        End If
    Next consumptionRow
    GeneratePivot = True
End Function

