Option Explicit

Sub 匯入週報並分析()

    Dim filePath As Variant

    Dim sourceWb As Workbook
    Dim targetWb As Workbook

    Dim ws As Worksheet
    Dim newWs As Worksheet
    Dim delWs As Worksheet
    Dim tempWs As Worksheet

    Dim lastRow As Long
    Dim lastCol As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set targetWb = ThisWorkbook

    filePath = Application.GetOpenFilename( _
        "Excel Files (*.xlsx), *.xlsx")

    If filePath = False Then
        MsgBox "未選擇檔案", vbExclamation
        GoTo ExitHandler
    End If

    Set tempWs = targetWb.Worksheets.Add
    tempWs.Name = "TEMP_IMPORT"

    For Each delWs In targetWb.Worksheets
        If delWs.Name <> "TEMP_IMPORT" Then
            delWs.Delete
        End If
    Next delWs

    Set sourceWb = Workbooks.Open(filePath, ReadOnly:=True)

    For Each ws In sourceWb.Worksheets

        If Left(ws.Name, 5) = "採購週報表" Then

            Set newWs = targetWb.Worksheets.Add( _
                After:=targetWb.Worksheets(targetWb.Worksheets.Count))

            On Error Resume Next
            newWs.Name = Left(ws.Name, 31)
            On Error GoTo 0

            If Not ws.Cells.Find("*", , , , xlByRows, xlPrevious) Is Nothing Then

                lastRow = ws.Cells.Find("*", , , , xlByRows, xlPrevious).Row
                lastCol = ws.Cells.Find("*", , , , xlByColumns, xlPrevious).Column

                newWs.Range( _
                    newWs.Cells(1, 1), _
                    newWs.Cells(lastRow, lastCol)).Value = _
                    ws.Range( _
                    ws.Cells(1, 1), _
                    ws.Cells(lastRow, lastCol)).Value

            End If

        End If

    Next ws

    sourceWb.Close False

    On Error Resume Next
    targetWb.Worksheets("TEMP_IMPORT").Delete
    On Error GoTo 0

    Call 開始缺貨分析

ExitHandler:

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "週報匯入與分析完成！", vbInformation

End Sub



