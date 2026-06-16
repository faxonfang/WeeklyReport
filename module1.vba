Option Explicit

Sub 開始缺貨分析()

    Dim ws As Worksheet
    Dim summaryWs As Worksheet

    Dim lastRow As Long, lastCol As Long, outRow As Long
    Dim i As Long, c As Long, x As Long

    Dim totalStock As Double
    Dim sales30 As Double, sales60 As Double, sales90 As Double, sales120 As Double
    Dim real30 As Double, real60 As Double, real90 As Double, real120 As Double
    Dim weightedSales As Double, dailySales As Double, stockDays As Double

    Dim demand60 As Double
    Dim incomingPO As Double
    Dim futureStock As Double
    Dim needOrder As String
    Dim suggestQty As Double

    Dim etdText As String, etdDisplay As String
    Dim etdDate As Date
    Dim etdDays As Long, minEtdDays As Long

    Dim poNo As String, poQty As Double
    Dim poList As String, poQtyList As String, etdList As String

    Dim riskLevel As String
    Dim arrDate
    Dim cleanText As String, ch As String
    Dim mm As Long, dd As Long
    Dim isLowStock As Boolean

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error Resume Next
    Worksheets("缺貨總覽").Delete
    On Error GoTo 0

    Set summaryWs = Worksheets.Add
    summaryWs.Name = "缺貨總覽"

    summaryWs.Range("A1:S1").Value = Array( _
        "工廠", "品號", "品名", "總庫存", _
        "30天銷售", "60天銷售", "90天銷售", "120天銷售", _
        "加權月銷", "可售天數", "60天需求", _
        "在途PO總量", "60天後預估庫存", _
        "是否建議下單", "建議下單量", _
        "PO號碼", "PO數量", "ETD時間", "風險等級")

    summaryWs.Columns("P:R").NumberFormat = "@"

    outRow = 2

    For Each ws In Worksheets

        If ws.Name <> "缺貨總覽" _
        And ws.Name <> "TEMP_IMPORT" _
        And Left(ws.Name, 5) = "採購週報表" Then

            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column

            For i = 4 To lastRow

                If Trim(ws.Cells(i, "A").Value) <> "" Then

                    If Trim(ws.Cells(i, "AA").Value) = "" _
                    Or ws.Cells(i, "AA").Value = 0 Then

                        totalStock = 取得最小上華庫存(ws, ws.Cells(i, "A").Value, lastRow)

                        sales30 = Val(ws.Cells(i, "N").Value)
                        sales60 = Val(ws.Cells(i, "O").Value)
                        sales90 = Val(ws.Cells(i, "P").Value)
                        sales120 = Val(ws.Cells(i, "Q").Value)

                        real30 = sales30
                        real60 = sales60 - sales30
                        real90 = sales90 - sales60
                        real120 = sales120 - sales90

                        If real60 < 0 Then real60 = 0
                        If real90 < 0 Then real90 = 0
                        If real120 < 0 Then real120 = 0

                        isLowStock = False

                        If totalStock <= real30 / 2 Then
                            isLowStock = True
                        End If

                        If isLowStock = True Then
                            weightedSales = _
                                real30 * 0.1 + _
                                real60 * 0.2 + _
                                real90 * 0.4 + _
                                real120 * 0.3
                        Else
                            weightedSales = _
                                real30 * 0.5 + _
                                real60 * 0.3 + _
                                real90 * 0.15 + _
                                real120 * 0.05
                        End If

                        If weightedSales > 0 Then

                            dailySales = weightedSales / 30
                            stockDays = totalStock / dailySales

                            poList = ""
                            poQtyList = ""
                            etdList = ""
                            incomingPO = 0
                            minEtdDays = 9999

                            For c = 1 To lastCol

                                If InStr(ws.Cells(2, c).Text, "ETD") > 0 _
                                Or InStr(ws.Cells(3, c).Text, "PO") > 0 Then

                                    poQty = Val(ws.Cells(i, c).Value)

                                    If poQty > 0 Then

                                        incomingPO = incomingPO + poQty

                                        poNo = ws.Cells(3, c).Text

                                        etdDisplay = Replace(ws.Cells(2, c).Text, vbLf, " ")
                                        etdDisplay = Replace(etdDisplay, vbCr, " ")

                                        If poList = "" Then
                                            poList = poNo
                                            poQtyList = CStr(poQty)
                                            etdList = etdDisplay
                                        Else
                                            poList = poList & " / " & poNo
                                            poQtyList = poQtyList & " / " & CStr(poQty)
                                            etdList = etdList & " / " & etdDisplay
                                        End If

                                        etdText = etdDisplay
                                        etdText = Replace(etdText, "ETD:", "")
                                        etdText = Replace(etdText, "ETD：", "")
                                        etdText = Replace(etdText, "-", "/")
                                        etdText = Trim(etdText)

                                        cleanText = ""

                                        For x = 1 To Len(etdText)

                                            ch = Mid(etdText, x, 1)

                                            If (ch >= "0" And ch <= "9") Or ch = "/" Then
                                                cleanText = cleanText & ch
                                            End If

                                        Next x

                                        If InStr(cleanText, "/") > 0 Then

                                            arrDate = Split(cleanText, "/")

                                            On Error Resume Next

                                            If UBound(arrDate) >= 1 Then

                                                If UBound(arrDate) >= 2 Then
                                                    mm = Val(arrDate(1))
                                                    dd = Val(arrDate(2))
                                                Else
                                                    mm = Val(arrDate(0))
                                                    dd = Val(arrDate(1))
                                                End If

                                                If mm >= 1 And mm <= 12 _
                                                And dd >= 1 And dd <= 31 Then

                                                    etdDate = DateSerial(Year(Date), mm, dd)
                                                    etdDays = DateDiff("d", Date, etdDate)

                                                    If etdDays > 0 Then
                                                        If etdDays < minEtdDays Then
                                                            minEtdDays = etdDays
                                                        End If
                                                    End If

                                                End If

                                            End If

                                            On Error GoTo 0

                                        End If

                                    End If

                                End If

                            Next c

                            If minEtdDays = 9999 Then
                                minEtdDays = 30
                            End If

                            demand60 = dailySales * 60
                            futureStock = totalStock + incomingPO - demand60

                            needOrder = "否"
                            suggestQty = 0

                            If futureStock < weightedSales Then
                                needOrder = "建議下單"
                                suggestQty = Round(weightedSales * 3 - futureStock, 0)

                                If suggestQty < 0 Then
                                    suggestQty = 0
                                End If
                            End If

                            riskLevel = ""

                            If stockDays < minEtdDays * 0.5 Then
                                riskLevel = "高風險"
                            ElseIf stockDays < minEtdDays Then
                                riskLevel = "中高風險"
                            ElseIf stockDays < minEtdDays * 1.3 Then
                                riskLevel = "注意補貨"
                            End If

                            If riskLevel <> "" Or needOrder = "建議下單" Then

                                summaryWs.Cells(outRow, 1).Value = ws.Name
                                summaryWs.Cells(outRow, 2).Value = ws.Cells(i, "A").Value
                                summaryWs.Cells(outRow, 3).Value = ws.Cells(i, "B").Value
                                summaryWs.Cells(outRow, 4).Value = totalStock

                                summaryWs.Cells(outRow, 5).Value = sales30
                                summaryWs.Cells(outRow, 6).Value = sales60
                                summaryWs.Cells(outRow, 7).Value = sales90
                                summaryWs.Cells(outRow, 8).Value = sales120

                                summaryWs.Cells(outRow, 9).Value = Round(weightedSales, 1)
                                summaryWs.Cells(outRow, 10).Value = Round(stockDays, 1)
                                summaryWs.Cells(outRow, 11).Value = Round(demand60, 1)
                                summaryWs.Cells(outRow, 12).Value = incomingPO
                                summaryWs.Cells(outRow, 13).Value = Round(futureStock, 1)

                                summaryWs.Cells(outRow, 14).Value = needOrder
                                summaryWs.Cells(outRow, 15).Value = suggestQty

                                summaryWs.Cells(outRow, 16).Value = "'" & poList
                                summaryWs.Cells(outRow, 17).Value = "'" & poQtyList
                                summaryWs.Cells(outRow, 18).Value = "'" & etdList
                                summaryWs.Cells(outRow, 19).Value = riskLevel

                                Select Case riskLevel
                                    Case "高風險"
                                        summaryWs.Rows(outRow).Interior.Color = RGB(255, 199, 206)
                                    Case "中高風險"
                                        summaryWs.Rows(outRow).Interior.Color = RGB(255, 235, 156)
                                    Case "注意補貨"
                                        summaryWs.Rows(outRow).Interior.Color = RGB(255, 242, 204)
                                End Select

                                If needOrder = "建議下單" And riskLevel = "" Then
                                    summaryWs.Rows(outRow).Interior.Color = RGB(221, 235, 247)
                                End If

                                outRow = outRow + 1

                            End If

                        End If

                    End If

                End If

            Next i

        End If

    Next ws

    summaryWs.Columns.AutoFit

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "缺貨分析完成！", vbInformation

End Sub


Function 取得最小上華庫存(ws As Worksheet, productCode As String, lastRow As Long) As Double

    Dim baseCode As String
    Dim currentCode As String
    Dim i As Long
    Dim stockValue As Double
    Dim minStock As Double
    Dim suffix As String
    Dim hasPart As Boolean

    minStock = 999999
    baseCode = Trim(productCode)
    hasPart = False

    If Len(baseCode) >= 2 Then

        suffix = Right(baseCode, 2)

        If Left(suffix, 1) = "-" _
        And Mid(suffix, 2, 1) Like "[A-Za-z]" Then
            baseCode = Left(baseCode, Len(baseCode) - 2)
        End If

    End If

    For i = 4 To lastRow

        currentCode = Trim(ws.Cells(i, "A").Value)

        If currentCode Like baseCode & "-[A-Za-z]" Then

            hasPart = True
            stockValue = Val(ws.Cells(i, "D").Value)

            If stockValue < minStock Then
                minStock = stockValue
            End If

        End If

    Next i

    If hasPart = False Then

        For i = 4 To lastRow

            currentCode = Trim(ws.Cells(i, "A").Value)

            If currentCode = baseCode Then
                minStock = Val(ws.Cells(i, "D").Value)
                Exit For
            End If

        Next i

    End If

    If minStock = 999999 Then
        minStock = 0
    End If

    取得最小上華庫存 = minStock

End Function



