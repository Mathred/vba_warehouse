Attribute VB_Name = "Main"
Sub InsertRmDictionary()
    Application.DisplayAlerts = False
    If (InsertSheet("���������� RM")) Then
        If (Utils.FSheetExists("���������� ��������")) Then
            If (repivot()) Then
                Call Utils.RMFormat
                Call recalculate
            End If
        End If
    End If
    Call updateUi
    Application.DisplayAlerts = True
    Sheets("Main").Activate
End Sub
Sub InsertConsumptionDictionary()
    Application.DisplayAlerts = False
    If (InsertSheet("���������� ��������")) Then
        If (repivot()) Then
            Call recalculate
        End If
    End If
    Call updateUi
    Application.DisplayAlerts = True
    Sheets("Main").Activate
End Sub
Sub InsertDppBap()
    If (InsertSheet("DPP", "DPP_BAP")) Then
        Call recalculate
    End If
    Call updateUi
    Sheets("Main").Activate
End Sub
Sub InsertDppNdc()
    If (InsertSheet("DPP", "DPP_NDC")) Then
        Call recalculate
    End If
    Call updateUi
    Sheets("Main").Activate
End Sub
Function InsertSheet(customerSheetName As String, Optional targetSheetName As String) As Boolean
    ' Get customer workbook...
    Dim filter As String
    Dim caption As String
    Dim customerFilename As String
    Dim customerWorkbook As Workbook
    Dim targetWorkbook As Workbook
    Application.DisplayAlerts = False
    ' make weak assumption that active workbook is the target
    Set targetWorkbook = Application.ActiveWorkbook
    
    ' get the customer workbook
    filter = "Excel Files (*.xlsx;*.xlsm),*.xlsx;*.xlsm"
    caption = "����������, �������� ���� � " & sheetName
    customerFilename = Application.GetOpenFilename(filter, , caption)
    If Not customerFilename = "False" Then
        Set customerWorkbook = Application.Workbooks.Open(customerFilename, 0)
        If Utils.FSheetExists(customerSheetName, customerWorkbook) Then
            If (Utils.FSheetExists(targetSheetName, targetWorkbook)) Then
                Application.DisplayAlerts = False
                targetWorkbook.Sheets(targetSheetName).Delete
                Application.DisplayAlerts = True
            End If
            If (Utils.FSheetExists(customerSheetName, targetWorkbook)) Then
                Application.DisplayAlerts = False
                targetWorkbook.Sheets(customerSheetName).Delete
                Application.DisplayAlerts = True
            End If
            customerWorkbook.Sheets(customerSheetName).Copy After:=targetWorkbook.Sheets(1)
                        
            If (customerSheetName = "DPP") Then
                If (targetSheetName = "DPP_BAP") Then
                    Call Utils.SDeleteOrdersSheets("BAP")
                    Call importOrdersIfExist(customerWorkbook, targetWorkbook, 1, "BAP")
                    Call importOrdersIfExist(customerWorkbook, targetWorkbook, 2, "BAP")
                ElseIf (targetSheetName = "DPP_NDC") Then
                    Call Utils.SDeleteOrdersSheets("NDC")
                    Call importOrdersIfExist(customerWorkbook, targetWorkbook, 1, "NDC")
                    Call importOrdersIfExist(customerWorkbook, targetWorkbook, 2, "NDC")
                End If
            End If
            If targetSheetName <> "" Then
                If (Utils.FSheetExists(targetSheetName, targetWorkbook)) Then
                    Application.DisplayAlerts = False
                    targetWorkbook.Sheets(targetSheetName).Delete
                    Application.DisplayAlerts = True
                End If
                targetWorkbook.Sheets(customerSheetName).Name = targetSheetName
            End If
        Else
            Application.DisplayAlerts = True
            MsgBox "���� " & customerSheetName & " �� ������", vbExclamation, "ERROR"
            InsertSheet = False
            customerWorkbook.Close
            Exit Function
        End If
        Application.DisplayAlerts = True
        customerWorkbook.Close (0)
        InsertSheet = True
    Else
        Application.DisplayAlerts = True
        MsgBox "���� �� ������", vbExclamation, "ERROR"
        InsertSheet = False
        Exit Function
    End If
    Sheets("Main").Activate
End Function
Sub importOrdersIfExist(customerWorkbook As Workbook, targetWorkbook As Workbook, weekNumber As Integer, Optional dpp As String)
    Dim orderSheetName As String: orderSheetName = Utils.getOrderSheetName(weekNumber)
    If Utils.FSheetExists(orderSheetName, customerWorkbook) Then
        If (Utils.FSheetExists(orderSheetName, targetWorkbook)) Then
            Application.DisplayAlerts = False
            targetWorkbook.Sheets(orderSheetName).Delete
            Application.DisplayAlerts = True
        End If
        customerWorkbook.Sheets(orderSheetName).Copy After:=targetWorkbook.Sheets(1)
        If dpp <> "" Then
            targetWorkbook.Sheets(orderSheetName).Name = orderSheetName & " " & dpp
        End If
    End If
    Sheets("Main").Activate
End Sub
Sub deleteDPPData()
    If MsgBox("��� ������ ������� DPP, ������ � ����������� �������! �� �������?", vbYesNo) = vbNo Then Exit Sub
    Dim i As Integer
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    For i = Sheets.Count To 1 Step -1
        Set ws = Sheets(i)
        If (ws.Name <> "Main") And (ws.Name <> "���������� ��������") And (ws.Name <> "���������� RM") And (ws.Name <> "Pivot") Then
            Sheets(i).Delete
        End If
    Next i
    Application.DisplayAlerts = True
    updateUi
    Sheets("Main").Activate
End Sub
Sub deleteTimetables()
    If MsgBox("��� ������ ����������� �������! �� �������?", vbYesNo) = vbNo Then Exit Sub
    Dim i As Integer
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    For i = Sheets.Count To 1 Step -1
        Set ws = Sheets(i)
        If (ws.Name <> "Main") And (ws.Name <> "���������� ��������") And (ws.Name <> "���������� RM") And (ws.Name <> "Pivot") And (ws.Name <> "DPP_BAP") And (ws.Name <> "DPP_NDC") And InStr(ws.Name, "�����") = 0 And (ws.Name <> "Records") Then
            Sheets(i).Delete
        End If
    Next i
    Sheets("Main").Activate
    Application.DisplayAlerts = True
    updateUi
End Sub
Function repivot() As Boolean
    If (Utils.IsPivotReadyToGenerate()) Then
        repivot = Pivot.GeneratePivot
    Else
        repivot = False
    End If
    Sheets("Main").Activate
End Function
Sub manualrepivot()
Call repivot
Sheets("Main").Activate
End Sub
Sub deleteAllData()
    If MsgBox("��� ������ ��� ����������� � �������! �� �������?", vbYesNo) = vbNo Then Exit Sub
    Dim i As Integer
    Application.DisplayAlerts = False
    For i = Sheets.Count To 1 Step -1
        If (Sheets(i).Name <> "Main") Then
            Sheets(i).Delete
        End If
    Next i
    Application.DisplayAlerts = True
    updateUi
    Sheets("Main").Activate
End Sub
Sub recalculate()
    If (Utils.IsDictionaryReadyToCalculate()) Then
        Dictionary.GenerateRecordsTable
        Calculation.CalculateTimeTables
    End If
    Sheets("Main").Activate
End Sub
Sub updateUi()
    Dim mainSheet As Worksheet: Set mainSheet = ActiveWorkbook.Sheets("Main")
    
    Call setStatus(mainSheet.Range("F2"), "���������� RM")
    Call setStatus(mainSheet.Range("F4"), "���������� ��������")
    Call setStatus(mainSheet.Range("F6"), "DPP_BAP")
    
    If (Utils.FSheetExists("DPP_BAP")) Then
        Call setStatus(mainSheet.Range("F8"), Utils.getOrderSheetName(1, , Sheets("DPP_BAP")) & " BAP")
        Call setStatus(mainSheet.Range("F10"), Utils.getOrderSheetName(2, , Sheets("DPP_BAP")) & " BAP")
    Else
        Call setStatus(mainSheet.Range("F8"), "ASDASDASDASD")
        Call setStatus(mainSheet.Range("F10"), "DFGDFGDFGDSF")
    End If
    
    Call setStatus(mainSheet.Range("F12"), "DPP_NDC")
    
    If (Utils.FSheetExists("DPP_NDC")) Then
        Call setStatus(mainSheet.Range("F14"), Utils.getOrderSheetName(1, , Sheets("DPP_NDC")) & " NDC")
        Call setStatus(mainSheet.Range("F16"), Utils.getOrderSheetName(2, , Sheets("DPP_BAP")) & " NDC")
    Else
        Call setStatus(mainSheet.Range("F14"), "ASDASDASDASD")
        Call setStatus(mainSheet.Range("F16"), "DFGDFGDFGDSF")
    End If
    
    Call setGenerateStatus(mainSheet.Range("H2"), "Pivot")
    Call setGenerateStatus(mainSheet.Range("H3"), "Records")
    If (Utils.IsAtLeastOneDateInSheets()) Then
        mainSheet.Range("H4") = "Time Table �������������"
        mainSheet.Range("H4").Interior.color = RGB(0, 300, 200)
    Else
        mainSheet.Range("H4") = "Time Table �� �������������"
        mainSheet.Range("H4").Interior.color = RGB(255, 0, 0)
    End If
    
    mainSheet.Activate
    Sheets("Main").Activate
End Sub
Sub setStatus(rng As Range, shtName As String)
    Dim noSheet As Long: noSheet = RGB(255, 0, 0)
    Dim yesSheet As Long: yesSheet = RGB(0, 300, 200)
    If (Utils.FSheetExists(shtName)) Then
        rng = "���������"
        rng.Interior.color = yesSheet
    Else
        rng = "�����������"
        rng.Interior.color = noSheet
    End If
    Sheets("Main").Activate
End Sub
Sub setGenerateStatus(rng As Range, shtName As String)
    Dim noSheet As Long: noSheet = RGB(255, 0, 0)
    Dim yesSheet As Long: yesSheet = RGB(0, 300, 200)
    If (Utils.FSheetExists(shtName)) Then
        rng = shtName & " ������������"
        rng.Interior.color = yesSheet
    Else
        rng = shtName & " �� ������������"
        rng.Interior.color = noSheet
    End If
    Sheets("Main").Activate
End Sub
Sub InsertFirstWeekOrderNDC()
    If (Utils.FSheetExists("DPP_NDC")) Then
        If (InsertSheet(Utils.getOrderSheetName(1, , Sheets("DPP_NDC"))) & " NDC") Then
            Call recalculate
        End If
    Else
        MsgBox ("������� �������� DPP")
    End If
    Call updateUi
    Sheets("Main").Activate
End Sub
Sub InsertSecondWeekOrderNDC()
    If (Utils.FSheetExists("DPP_NDC")) Then
        If (InsertSheet(Utils.getOrderSheetName(2, , Sheets("DPP_NDC"))) & " NDC") Then
            Call recalculate
        End If
    Else
        MsgBox ("������� �������� DPP")
    End If
    Call updateUi
    Sheets("Main").Activate
End Sub
Sub InsertFirstWeekOrderBAP()
    If (Utils.FSheetExists("DPP_BAP")) Then
        If (InsertSheet(Utils.getOrderSheetName(1, , Sheets("DPP_BAP"))) & " BAP") Then
            Call recalculate
        End If
    Else
        MsgBox ("������� �������� DPP")
    End If
    Call updateUi
    Sheets("Main").Activate

End Sub

Sub InsertSecondWeekOrderBAP()
    If (Utils.FSheetExists("DPP_BAP")) Then
        If (InsertSheet(Utils.getOrderSheetName(2, , Sheets("DPP_BAP"))) & " BAP") Then
            Call recalculate
        End If
    Else
        MsgBox ("������� �������� DPP")
    End If
    Call updateUi
    Sheets("Main").Activate
End Sub
Sub manualRecalculate()
    Dictionary.GenerateRecordsTable
    Calculation.CalculateTimeTables
    Call updateUi
    Sheets("Main").Activate
End Sub
Sub allTables()
    Dim TablesUtils As New TablesUtils
    Call Tables.CreateFGTable
    Call TablesUtils.MergeTables("FG")
    Call Tables.CreateIDTable
    Call TablesUtils.MergeTables("ID")
    Call Calculation.CalculateTimeTables
    
    Sheets("Main").Activate
    Call updateUi
End Sub
