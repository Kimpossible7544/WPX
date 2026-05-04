Option Explicit

'=======================================================
' DASHBOARD - Build the Player Dashboard sheet
'=======================================================
Public Sub BuildPlayerDashboard()

    Dim wsDash As Worksheet
    Dim wsPT As Worksheet
    Dim firstPlayer As String

    Set wsPT = ThisWorkbook.Worksheets("Player Tracking")
    Set wsDash = GetOrCreateSheet("Player Dashboard")

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo ErrHandler

    StyleDashboardSheet wsDash
    BuildPlayerDropdownFromTracking wsDash, wsPT
    BuildDashboardCharts wsDash

    firstPlayer = Trim(SafeStr(wsDash.Range("J2").Value))
    If firstPlayer <> "" Then
        wsDash.Range("B3").Value = firstPlayer
        RefreshDashboard
    End If

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "BuildPlayerDashboard error " & Err.Number & ": " & Err.Description, vbCritical
    Resume SafeExit

End Sub

Public Sub UpdateWinLoss()

    Dim ws As Worksheet
    Dim result As String
    Dim val As String

    result = ""

    For Each ws In ThisWorkbook.Worksheets

        If ws.Name <> "Roster" _
        And ws.Name <> "Score Import" _
        And ws.Name <> "Player Tracking" _
        And ws.Name <> "Archived Players" _
        And ws.Name <> "Player Dashboard" Then

            val = Trim(SafeStr(ws.Range("A120").Value))

            If val <> "" Then
                If UCase(val) = "WPX" Then
                    result = result & "W "
                Else
                    result = result & "L "
                End If
            End If

        End If

    Next ws

    ThisWorkbook.Worksheets("Player Tracking").Range("F2").Value = Trim(result)

End Sub

'=======================================================
' Optional hard reset / rebuild
'=======================================================
Public Sub DashboardHardReset()

    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Worksheets("Player Dashboard")

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo SafeExit

    wsDash.Cells.Clear
    BuildPlayerDashboard

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

'=======================================================
' Style the dashboard sheet
'=======================================================
Public Sub StyleDashboardSheet(wsDash As Worksheet)

    Dim bgColor As Long
    Dim accentColor As Long
    Dim win As Window

    bgColor = RGB(15, 15, 25)
    accentColor = RGB(0, 255, 180)

    wsDash.Cells.Clear
    wsDash.Cells.Interior.Color = bgColor
    wsDash.Cells.Font.Color = RGB(220, 220, 220)

    For Each win In ThisWorkbook.Windows
        win.DisplayGridlines = False
    Next win

    wsDash.Columns("A").ColumnWidth = 2
    wsDash.Columns("B").ColumnWidth = 18
    wsDash.Columns("C").ColumnWidth = 22
    wsDash.Columns("D").ColumnWidth = 22
    wsDash.Columns("E").ColumnWidth = 22
    wsDash.Columns("F").ColumnWidth = 16
    wsDash.Columns("G").ColumnWidth = 16
    wsDash.Columns("H").ColumnWidth = 18
    wsDash.Columns("I").ColumnWidth = 2
    wsDash.Columns("J").ColumnWidth = 16
    wsDash.Columns("K").ColumnWidth = 16
    wsDash.Columns("L").ColumnWidth = 16

    wsDash.Rows("1").RowHeight = 8
    wsDash.Rows("2").RowHeight = 50
    wsDash.Rows("3").RowHeight = 35
    wsDash.Rows("4").RowHeight = 8
    wsDash.Rows("5").RowHeight = 25
    wsDash.Rows("6").RowHeight = 40
    wsDash.Rows("8").RowHeight = 28
    wsDash.Rows("9").RowHeight = 28

    With wsDash.Range("B2:E2")
        .UnMerge
        .Merge
        .Value = "WARRIOR PHOENIX"
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = accentColor
        .Font.Name = "Segoe UI"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(20, 20, 35)
        .Borders(xlEdgeBottom).Color = accentColor
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With

    With wsDash.Range("F2:G2")
        .UnMerge
        .Merge
        .Value = ""
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = accentColor
        .Font.Name = "Segoe UI"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(20, 20, 35)
        .Borders(xlEdgeBottom).Color = accentColor
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With

    With wsDash.Range("B3")
        .Value = ""
        .Interior.Color = RGB(25, 25, 40)
        .Font.Color = RGB(220, 220, 220)
        .Font.Size = 13
        .Font.Name = "Segoe UI"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).Color = accentColor
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeLeft).Color = accentColor
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeRight).Color = accentColor
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeTop).Color = accentColor
        .Borders(xlEdgeTop).Weight = xlMedium
    End With

    With wsDash.Range("C3:F3")
        .UnMerge
        .Merge
        .Value = ""
        .Font.Size = 30
        .Font.Bold = True
        .Font.Name = "Informal Roman"
        .Font.Color = RGB(6, 255, 172)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(10, 10, 20)
    End With

    On Error Resume Next
    wsDash.Range("C3").Font.Shadow = True
    On Error GoTo 0

    StyleStatBox wsDash, "B5", "OVERALL RANK", RGB(0, 255, 180)
    StyleStatBox wsDash, "C5", "TOTAL SCORE", RGB(0, 180, 255)
    StyleStatBox wsDash, "D5", "WEEKLY AVG", RGB(180, 100, 255)
    StyleStatBox wsDash, "E5", "WEEKS PLAYED", RGB(255, 200, 0)
    StyleStatBox wsDash, "F5", "MISSED DAILY", RGB(255, 100, 50)
    StyleStatBox wsDash, "G5", "MISSED WEEKLY", RGB(255, 50, 150)
    StyleStatBox wsDash, "H5", "NUMBER OF DNP", RGB(0, 255, 180)

    StyleStatValue wsDash, "B6", RGB(0, 255, 180)
    StyleStatValue wsDash, "C6", RGB(0, 180, 255)
    StyleStatValue wsDash, "D6", RGB(180, 100, 255)
    StyleStatValue wsDash, "E6", RGB(255, 200, 0)
    StyleStatValue wsDash, "F6", RGB(255, 100, 50)
    StyleStatValue wsDash, "G6", RGB(255, 50, 150)
    StyleStatValue wsDash, "H6", RGB(0, 255, 180)

    wsDash.Range("J1").Value = "PlayerList"
    wsDash.Columns("J:L").Hidden = True

End Sub

Private Sub StyleStatBox(ws As Worksheet, cellAddr As String, labelText As String, boxColor As Long)
    With ws.Range(cellAddr)
        .Value = labelText
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = boxColor
        .Font.Name = "Segoe UI"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(20, 20, 35)
        .Borders(xlEdgeTop).Color = boxColor
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeLeft).Color = RGB(40, 40, 60)
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeRight).Color = RGB(40, 40, 60)
        .Borders(xlEdgeRight).Weight = xlThin
    End With
End Sub

Private Sub StyleStatValue(ws As Worksheet, cellAddr As String, boxColor As Long)
    With ws.Range(cellAddr)
        .Value = "-"
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = boxColor
        .Font.Name = "Segoe UI"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(20, 20, 35)
        .Borders(xlEdgeBottom).Color = boxColor
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeLeft).Color = RGB(40, 40, 60)
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeRight).Color = RGB(40, 40, 60)
        .Borders(xlEdgeRight).Weight = xlThin
    End With
End Sub

'=======================================================
' Build player dropdown in B3 from Player Tracking
'=======================================================
Private Sub BuildPlayerDropdownFromTracking(wsDash As Worksheet, wsPT As Worksheet)

    Dim lastRow As Long
    Dim destRow As Long
    Dim i As Long, j As Long
    Dim nm As String
    Dim temp As String
    Dim namesArr() As String
    Dim countNames As Long

    lastRow = GetLastRow(wsPT, 1)
    If lastRow < 2 Then Exit Sub

    wsDash.Range("J2:J500").ClearContents
    countNames = 0

    For i = 2 To lastRow
        nm = Trim(SafeStr(wsPT.Cells(i, "A").Value))
        If nm <> "" Then
            countNames = countNames + 1
            ReDim Preserve namesArr(1 To countNames)
            namesArr(countNames) = nm
        End If
    Next i

    If countNames = 0 Then
        wsDash.Range("B3").Validation.Delete
        Exit Sub
    End If

    For i = 1 To countNames - 1
        For j = i + 1 To countNames
            If UCase(namesArr(j)) < UCase(namesArr(i)) Then
                temp = namesArr(i)
                namesArr(i) = namesArr(j)
                namesArr(j) = temp
            End If
        Next j
    Next i

    destRow = 2
    For i = 1 To countNames
        wsDash.Cells(destRow, "J").Value = namesArr(i)
        destRow = destRow + 1
    Next i

    With wsDash.Range("B3").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="=$J$2:$J$" & destRow - 1
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = False
    End With

End Sub

'=======================================================
' Build the charts on the dashboard
'=======================================================
Private Sub BuildDashboardCharts(wsDash As Worksheet)

    Dim co As ChartObject
    Dim cht1 As ChartObject
    Dim cht2 As ChartObject
    Dim cht3 As ChartObject

    For Each co In wsDash.ChartObjects
        co.Delete
    Next co

    Set cht1 = wsDash.ChartObjects.Add(Left:=10, Top:=180, Width:=480, Height:=260)
    cht1.Name = "ScoreChart"

    With cht1.Chart
        .ChartType = xlLineMarkers
        .HasTitle = True
        .ChartTitle.Text = "WEEKLY SCORE TREND"
        .ChartTitle.Font.Color = RGB(0, 255, 180)
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Name = "Segoe UI"
        .PlotArea.Interior.Color = RGB(20, 20, 35)
        .PlotArea.Border.Color = RGB(40, 40, 60)
        .ChartArea.Interior.Color = RGB(15, 15, 25)
        .ChartArea.Border.Color = RGB(40, 40, 60)
        .HasLegend = True
        .Legend.Font.Color = RGB(180, 180, 200)
        .Legend.Font.Size = 9
    End With

    Set cht2 = wsDash.ChartObjects.Add(Left:=500, Top:=180, Width:=480, Height:=260)
    cht2.Name = "MissedGoalsChart"

    With cht2.Chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "MISSED GOALS BY WEEK"
        .ChartTitle.Font.Color = RGB(255, 50, 150)
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Name = "Segoe UI"
        .PlotArea.Interior.Color = RGB(20, 20, 35)
        .PlotArea.Border.Color = RGB(40, 40, 60)
        .ChartArea.Interior.Color = RGB(15, 15, 25)
        .ChartArea.Border.Color = RGB(40, 40, 60)
        .HasLegend = True
        .Legend.Font.Color = RGB(180, 180, 200)
        .Legend.Font.Size = 9
    End With

    Set cht3 = wsDash.ChartObjects.Add(Left:=10, Top:=460, Width:=970, Height:=260)
    cht3.Name = "RankChart"

    With cht3.Chart
        .ChartType = xlLineMarkers
        .HasTitle = True
        .ChartTitle.Text = "WEEKLY RANK TREND"
        .ChartTitle.Font.Color = RGB(255, 200, 0)
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Name = "Segoe UI"
        .PlotArea.Interior.Color = RGB(20, 20, 35)
        .PlotArea.Border.Color = RGB(40, 40, 60)
        .ChartArea.Interior.Color = RGB(15, 15, 25)
        .ChartArea.Border.Color = RGB(40, 40, 60)
        .HasLegend = True
        .Legend.Font.Color = RGB(180, 180, 200)
        .Legend.Font.Size = 9
    End With

End Sub

'=======================================================
' Refresh dashboard when player is selected/opened
'=======================================================
Public Sub RefreshDashboard()

    Dim wsDash As Worksheet
    Dim wsPT As Worksheet
    Dim ws As Worksheet
    Dim player As String
    Dim playerID As String
    Dim ptRow As Long
    Dim lastCol As Long
    Dim col As Long
    Dim hdr As String
    Dim weekLabel As String
    Dim weekEndDate As Date
    Dim weekParts() As String
    Dim weekEndStr As String

    Dim weekNames() As Variant
    Dim weekScores() As Variant
    Dim weekRanks() As Variant
    Dim trendLineVals() As Variant
    Dim goalLineVals() As Variant
    Dim weekMissedDaily() As Variant
    Dim weekMissedWeekly() As Variant
    Dim weekWins() As Variant
    Dim weekLosses() As Variant
    Dim weekCount As Long

    Dim runningTotal As Double
    Dim totalWins As Long
    Dim totalLosses As Long
    Dim totalMissedDaily As Long
    Dim totalMissedWeekly As Long
    Dim weeksPlayed As Long
    Dim dnpCount As Long

    Dim i As Long
    Dim rw As Long
    Dim foundSheet As Boolean
    Dim weekSheetName As String
    Dim foundCell As Range
    Dim wl As Variant
    Dim co As ChartObject
    Dim oldCalc As XlCalculation

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    oldCalc = Application.Calculation
    Application.Calculation = xlCalculationManual

    Set wsDash = ThisWorkbook.Worksheets("Player Dashboard")
    Set wsPT = ThisWorkbook.Worksheets("Player Tracking")

    player = Trim(Replace(SafeStr(wsDash.Range("B3").Value), Chr(160), " "))
    If player = "" Then GoTo SafeExit

    wsDash.Range("C3").Value = player

    ptRow = 0
    For Each foundCell In wsPT.Range("A2:A" & GetLastRow(wsPT, 1))
        If Not IsError(foundCell.Value) Then
            If StrComp(Trim(Replace(CStr(foundCell.Value), Chr(160), " ")), player, vbTextCompare) = 0 Then
                ptRow = foundCell.Row
                Exit For
            End If
        End If
    Next foundCell

    If ptRow = 0 Then
        wsDash.Range("B6:H6").Value = "-"
        wsDash.Range("F2:G2").Value = ""
        ClearAllDashboardCharts wsDash
        GoTo SafeExit
    End If

    playerID = Trim(SafeStr(wsPT.Cells(ptRow, "CK").Value))

    lastCol = wsPT.Cells(1, wsPT.Columns.Count).End(xlToLeft).Column
    weekCount = 0

    ReDim weekNames(1 To 200)
    ReDim weekScores(1 To 200)
    ReDim weekRanks(1 To 200)
    ReDim trendLineVals(1 To 200)
    ReDim goalLineVals(1 To 200)
    ReDim weekMissedDaily(1 To 200)
    ReDim weekMissedWeekly(1 To 200)
    ReDim weekWins(1 To 200)
    ReDim weekLosses(1 To 200)

    For col = 7 To lastCol Step 2
        hdr = SafeStr(wsPT.Cells(1, col).Value)

        If InStr(1, hdr, "Weekly Total (", vbTextCompare) > 0 Then
            weekLabel = Replace(Replace(hdr, "Weekly Total (", "", , , vbTextCompare), ")", "")

            If InStr(weekLabel, " - ") > 0 Then
                weekParts = Split(weekLabel, " - ")
                weekEndStr = Trim(weekParts(UBound(weekParts)))

                weekEndDate = 0
                On Error Resume Next
                weekEndDate = CDate(weekEndStr & " " & Year(Now()))
                On Error GoTo ErrHandler

                If weekEndDate > Date + 7 Then
                    weekEndDate = DateSerial(Year(Now()) - 1, Month(weekEndDate), Day(weekEndDate))
                End If
                If weekEndDate >= Date Then GoTo NextWeekCol
            End If

            weekCount = weekCount + 1
            weekNames(weekCount) = weekLabel

            If Not IsError(wsPT.Cells(ptRow, col).Value) Then
                If IsNumeric(wsPT.Cells(ptRow, col).Value) Then
                    weekScores(weekCount) = CDbl(wsPT.Cells(ptRow, col).Value)
                Else
                    weekScores(weekCount) = 0
                End If
            Else
                weekScores(weekCount) = 0
            End If

            If col + 1 <= lastCol Then
                If Not IsError(wsPT.Cells(ptRow, col + 1).Value) Then
                    If IsNumeric(wsPT.Cells(ptRow, col + 1).Value) Then
                        weekRanks(weekCount) = CLng(wsPT.Cells(ptRow, col + 1).Value)
                    Else
                        weekRanks(weekCount) = Empty
                    End If
                Else
                    weekRanks(weekCount) = Empty
                End If
            End If
        End If

NextWeekCol:
    Next col

    If weekCount = 0 Then
        wsDash.Range("B6").Value = SafeStr(wsPT.Cells(ptRow, "C").Value)
        wsDash.Range("C6").Value = SafeFormat(wsPT.Cells(ptRow, "B").Value, "#,##0")
        wsDash.Range("D6").Value = SafeFormat(wsPT.Cells(ptRow, "D").Value, "#,##0")
        wsDash.Range("E6").Value = 0
        wsDash.Range("F6").Value = SafeStr(wsPT.Cells(ptRow, "E").Value)
        wsDash.Range("G6").Value = SafeStr(wsPT.Cells(ptRow, "F").Value)
        wsDash.Range("H6").Value = 0
        wsDash.Range("F2:G2").Value = ""
        ClearAllDashboardCharts wsDash
        GoTo SafeExit
    End If

    dnpCount = 0
    totalWins = 0
    totalLosses = 0
    totalMissedDaily = 0
    totalMissedWeekly = 0
    weeksPlayed = 0
    runningTotal = 0

    For i = 1 To weekCount

        runningTotal = runningTotal + CDbl(weekScores(i))
        trendLineVals(i) = runningTotal / i
        goalLineVals(i) = 20000000

        If CDbl(weekScores(i)) > 0 Then
            weeksPlayed = weeksPlayed + 1
        End If

        weekMissedDaily(i) = 0
        weekMissedWeekly(i) = 0
        weekWins(i) = 0
        weekLosses(i) = 0

        weekSheetName = CStr(weekNames(i))
        foundSheet = False

        For Each ws In ThisWorkbook.Worksheets
            If NormalizeDashboardName(ws.Name) = NormalizeDashboardName(weekSheetName) Then
                foundSheet = True

                Dim rowMatch As Boolean
                For rw = 2 To 106
                    rowMatch = False

                    If playerID <> "" Then
                        If Trim(SafeStr(ws.Cells(rw, "V").Value)) = playerID Then
                            rowMatch = True
                        End If
                    Else
                        If NormalizeDashboardName(SafeStr(ws.Cells(rw, "A").Value)) = NormalizeDashboardName(player) Then
                            rowMatch = True
                        End If
                    End If

                    If rowMatch Then
                        weekMissedDaily(i) = CountMissedDailyGoals(ws, rw)

                        If WeekIsComplete(ws, rw) Then
                            If Not IsError(ws.Cells(rw, "L").Value) Then
                                If IsNumeric(ws.Cells(rw, "L").Value) Then
                                    If CDbl(ws.Cells(rw, "L").Value) < 20000000 Then
                                        weekMissedWeekly(i) = 1
                                    End If
                                End If
                            End If
                        End If

                        dnpCount = dnpCount + CountPeachFailInRow(ws, rw)
                        Exit For
                    End If

                Next rw

                wl = GetWeekWinLossFromResultRow(ws)
                weekWins(i) = wl(0)
                weekLosses(i) = wl(1)

                Exit For
            End If
        Next ws

        totalMissedDaily = totalMissedDaily + CLng(weekMissedDaily(i))
        totalMissedWeekly = totalMissedWeekly + CLng(weekMissedWeekly(i))
        totalWins = totalWins + CLng(weekWins(i))
        totalLosses = totalLosses + CLng(weekLosses(i))
    Next i

    wsDash.Range("B6").Value = SafeStr(wsPT.Cells(ptRow, "C").Value)
    wsDash.Range("C6").Value = SafeFormat(wsPT.Cells(ptRow, "B").Value, "#,##0")
    wsDash.Range("D6").Value = SafeFormat(wsPT.Cells(ptRow, "D").Value, "#,##0")
    wsDash.Range("E6").Value = weeksPlayed
    wsDash.Range("F6").Value = totalMissedDaily
    wsDash.Range("G6").Value = totalMissedWeekly
    wsDash.Range("H6").Value = dnpCount

    With wsDash.Range("F2:G2")
        .UnMerge
        .Merge
        .Value = "W" & totalWins & " - L" & totalLosses
        .Font.Name = "Segoe UI"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(0, 255, 180)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    UpdateScoreChart wsDash, weekNames, weekScores, trendLineVals, goalLineVals, weekCount
    UpdateMissedGoalsChart wsDash, weekNames, weekMissedDaily, weekMissedWeekly, weekCount
    UpdateRankChartSimple wsDash, weekNames, weekRanks, weekCount

    On Error Resume Next
    wsDash.ChartObjects("ScoreChart").Chart.HasTitle = True
    wsDash.ChartObjects("ScoreChart").Chart.ChartTitle.Text = player & " - Weekly Scores"
    wsDash.ChartObjects("MissedGoalsChart").Chart.HasTitle = True
    wsDash.ChartObjects("MissedGoalsChart").Chart.ChartTitle.Text = player & " - Missed Goals"
    wsDash.ChartObjects("RankChart").Chart.HasTitle = True
    wsDash.ChartObjects("RankChart").Chart.ChartTitle.Text = player & " - Weekly Rank"
    On Error GoTo ErrHandler

    For Each co In wsDash.ChartObjects
        co.Placement = xlFreeFloating
        co.Chart.Refresh
        co.Visible = False
        co.Visible = True
    Next co

SafeExit:
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    DoEvents
    Exit Sub

ErrHandler:
    MsgBox "Dashboard Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume SafeExit

End Sub

Private Sub UpdateScoreChart(wsDash As Worksheet, weekNames As Variant, weekScores As Variant, trendLineVals As Variant, goalLineVals As Variant, weekCount As Long)

    Dim cht As Chart
    Dim s As Series
    Dim i As Long
    Dim startIx As Long
    Dim cnt As Long
    Dim plotNames() As Variant
    Dim plotScores() As Variant
    Dim plotTrend() As Variant
    Dim plotGoal() As Variant

    Set cht = wsDash.ChartObjects("ScoreChart").Chart

    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop

    If weekCount = 0 Then Exit Sub

    startIx = IIf(weekCount > 6, weekCount - 5, 1)
    cnt = weekCount - startIx + 1

    ReDim plotNames(1 To cnt)
    ReDim plotScores(1 To cnt)
    ReDim plotTrend(1 To cnt)
    ReDim plotGoal(1 To cnt)

    For i = 1 To cnt
        plotNames(i) = weekNames(startIx + i - 1)
        plotScores(i) = weekScores(startIx + i - 1)
        plotTrend(i) = trendLineVals(startIx + i - 1)
        plotGoal(i) = goalLineVals(startIx + i - 1)
    Next i

    Set s = cht.SeriesCollection.NewSeries
    With s
        .Name = "Weekly Score"
        .Values = plotScores
        .XValues = plotNames
        .Format.Line.ForeColor.RGB = RGB(0, 255, 180)
        .Format.Line.Weight = 2.5
        .MarkerStyle = xlMarkerStyleCircle
        .MarkerSize = 7
        .MarkerForegroundColor = RGB(0, 255, 180)
        .MarkerBackgroundColor = RGB(0, 255, 180)
        .HasDataLabels = True
        .DataLabels.Font.Color = RGB(0, 255, 180)
        .DataLabels.Font.Size = 7
        .DataLabels.NumberFormat = "#,##0"
        .DataLabels.Position = xlLabelPositionAbove
    End With

    Set s = cht.SeriesCollection.NewSeries
    With s
        .Name = "Trend Line"
        .Values = plotTrend
        .XValues = plotNames
        .Format.Line.ForeColor.RGB = RGB(255, 200, 0)
        .Format.Line.Weight = 2
        .MarkerStyle = xlMarkerStyleNone
    End With

    Set s = cht.SeriesCollection.NewSeries
    With s
        .Name = "20M Goal"
        .Values = plotGoal
        .XValues = plotNames
        .Format.Line.ForeColor.RGB = RGB(255, 50, 150)
        .Format.Line.Weight = 1.75
        .Format.Line.DashStyle = msoLineDash
        .MarkerStyle = xlMarkerStyleNone
    End With

End Sub

Private Sub UpdateMissedGoalsChart(wsDash As Worksheet, weekNames As Variant, weekMissedDaily As Variant, weekMissedWeekly As Variant, weekCount As Long)

    Dim cht As Chart
    Dim s As Series
    Dim i As Long
    Dim startIx As Long
    Dim cnt As Long
    Dim plotNames() As Variant
    Dim plotMissedDaily() As Variant
    Dim plotMissedWeekly() As Variant

    Set cht = wsDash.ChartObjects("MissedGoalsChart").Chart

    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop

    If weekCount = 0 Then Exit Sub

    startIx = IIf(weekCount > 6, weekCount - 5, 1)
    cnt = weekCount - startIx + 1

    ReDim plotNames(1 To cnt)
    ReDim plotMissedDaily(1 To cnt)
    ReDim plotMissedWeekly(1 To cnt)

    For i = 1 To cnt
        plotNames(i) = weekNames(startIx + i - 1)
        plotMissedDaily(i) = weekMissedDaily(startIx + i - 1)
        plotMissedWeekly(i) = weekMissedWeekly(startIx + i - 1)
    Next i

    Set s = cht.SeriesCollection.NewSeries
    With s
        .Name = "Missed Daily"
        .Values = plotMissedDaily
        .XValues = plotNames
        .Format.Fill.ForeColor.RGB = RGB(255, 100, 50)
        .Format.Line.Visible = msoFalse
    End With

    Set s = cht.SeriesCollection.NewSeries
    With s
        .Name = "Missed Weekly Goal"
        .Values = plotMissedWeekly
        .XValues = plotNames
        .Format.Fill.ForeColor.RGB = RGB(255, 50, 150)
        .Format.Line.Visible = msoFalse
    End With

End Sub

Private Sub UpdateRankChartSimple(wsDash As Worksheet, weekNames As Variant, weekRanks As Variant, weekCount As Long)

    Dim cht As Chart
    Dim s As Series
    Dim i As Long
    Dim plotNames() As Variant
    Dim plotRanks() As Variant
    Dim cnt As Long

    Set cht = wsDash.ChartObjects("RankChart").Chart

    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop

    cnt = 0
    ReDim plotNames(1 To weekCount)
    ReDim plotRanks(1 To weekCount)

    For i = 1 To weekCount
        If IsNumeric(weekRanks(i)) Then
            If CLng(weekRanks(i)) > 0 Then
                cnt = cnt + 1
                plotNames(cnt) = weekNames(i)
                plotRanks(cnt) = CLng(weekRanks(i))
            End If
        End If
    Next i

    If cnt = 0 Then Exit Sub

    ReDim Preserve plotNames(1 To cnt)
    ReDim Preserve plotRanks(1 To cnt)

    Set s = cht.SeriesCollection.NewSeries
    With s
        .Name = "Weekly Rank"
        .Values = plotRanks
        .XValues = plotNames
        .Format.Line.ForeColor.RGB = RGB(255, 200, 0)
        .Format.Line.Weight = 2.5
        .MarkerStyle = xlMarkerStyleDiamond
        .MarkerSize = 8
        .MarkerForegroundColor = RGB(255, 200, 0)
        .MarkerBackgroundColor = RGB(255, 200, 0)
        .HasDataLabels = True
        .DataLabels.Font.Color = RGB(255, 200, 0)
        .DataLabels.Font.Size = 7
        .DataLabels.Position = xlLabelPositionAbove
    End With

    On Error Resume Next
    cht.Axes(xlValue).ReversePlotOrder = True
    On Error GoTo 0

End Sub

'=======================================================
' Helpers
'=======================================================
Private Function GetOrCreateSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function

Private Function GetLastRow(ws As Worksheet, colNum As Long) As Long
    GetLastRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).Row
End Function

Private Function NormalizeDashboardName(s As String) As String
    NormalizeDashboardName = UCase(Trim(Replace(s, Chr(160), " ")))
End Function

Private Function CountMissedDailyGoals(ws As Worksheet, playerRow As Long) As Long
    Dim c As Range
    Dim cnt As Long
    cnt = 0
    For Each c In ws.Range("C" & playerRow & ":G" & playerRow)
        If Not IsError(c.Value) Then
            If IsNumeric(c.Value) Then
                If CDbl(c.Value) < 6000000 Then cnt = cnt + 1
            End If
        End If
    Next c
    CountMissedDailyGoals = cnt
End Function

Private Function ContainsPeachValue(v As Variant) As Boolean
    If IsError(v) Then
        ContainsPeachValue = False
        Exit Function
    End If
    Dim txt As String
    txt = Trim(CStr(v))
    If Len(txt) = 0 Then
        ContainsPeachValue = False
        Exit Function
    End If
    ContainsPeachValue = (LCase(txt) = "f") Or (LCase(txt) = "x") Or (AscW(txt) = 9746)
End Function

Private Function CountPeachFailInRow(ws As Worksheet, playerRow As Long) As Long
    Dim cnt As Long
    cnt = 0
    If ContainsPeachValue(ws.Cells(playerRow, "I").Value) Then cnt = cnt + 1
    If ContainsPeachValue(ws.Cells(playerRow, "J").Value) Then cnt = cnt + 1
    If ContainsPeachValue(ws.Cells(playerRow, "K").Value) Then cnt = cnt + 1
    CountPeachFailInRow = cnt
End Function

Private Function WeekIsComplete(ws As Worksheet, playerRow As Long) As Boolean
    WeekIsComplete = (Application.WorksheetFunction.CountA(ws.Range("C" & playerRow & ":G" & playerRow)) = 5)
End Function

Private Function GetWeekWinLossFromResultRow(ws As Worksheet) As Variant
    Dim result(1) As Long
    Dim txt As String
    result(0) = 0
    result(1) = 0

    txt = Trim(SafeStr(ws.Range("A120").Value))

    If txt = "" Then
        GetWeekWinLossFromResultRow = result
        Exit Function
    End If

    If UCase(txt) = "WPX" Then
        result(0) = 1
    Else
        result(1) = 1
    End If

    GetWeekWinLossFromResultRow = result
End Function

Private Sub ClearAllDashboardCharts(wsDash As Worksheet)
    On Error Resume Next
    Do While wsDash.ChartObjects.Count > 0
        wsDash.ChartObjects(1).Delete
    Loop
    BuildDashboardCharts wsDash
    On Error GoTo 0
End Sub

Public Sub OpenPlayerDashboard(playerName As String)
    Dim wsDash As Worksheet
    On Error Resume Next
    Set wsDash = ThisWorkbook.Worksheets("Player Dashboard")
    On Error GoTo 0
    If wsDash Is Nothing Then
        BuildPlayerDashboard
        Set wsDash = ThisWorkbook.Worksheets("Player Dashboard")
    End If
    wsDash.Range("B3").Value = playerName
    RefreshDashboard
    wsDash.Activate
End Sub

'=======================================================
' GetScore - Use in weekly sheet Col G: =GetScore(A2)
'=======================================================
Function GetScore(rawName As Variant) As Variant
    If IsError(rawName) Or IsEmpty(rawName) Then
        GetScore = ""
        Exit Function
    End If
    If rawName = "" Then
        GetScore = ""
        Exit Function
    End If

    Dim cleanName As String
    cleanName = FindRosterName(CStr(rawName))

    If cleanName = "" Then
        GetScore = ""
        Exit Function
    End If

    Dim wsImport As Worksheet
    On Error Resume Next
    Set wsImport = ThisWorkbook.Sheets("Score Import")
    On Error GoTo 0

    If wsImport Is Nothing Then
        GetScore = "NO IMPORT SHEET"
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = wsImport.Cells(wsImport.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow
        If Trim(LCase(SafeStr(wsImport.Cells(i, 1).Value))) = Trim(LCase(cleanName)) Then
            GetScore = wsImport.Cells(i, 2).Value
            Exit Function
        End If
    Next i

    GetScore = ""
End Function

'=======================================================
' FindRosterName
'=======================================================
Private Function FindRosterName(rawName As String) As String
    Dim wsRoster As Worksheet
    On Error Resume Next
    Set wsRoster = ThisWorkbook.Sheets("Roster")
    On Error GoTo 0

    If wsRoster Is Nothing Then
        FindRosterName = ""
        Exit Function
    End If

    Dim normRaw As String
    normRaw = NormalizeName(rawName)

    Dim nameCols(3) As Integer
    Dim akaCols(3) As Integer
    nameCols(0) = 2:  akaCols(0) = 4
    nameCols(1) = 6:  akaCols(1) = 8
    nameCols(2) = 10: akaCols(2) = 12
    nameCols(3) = 14: akaCols(3) = 16

    Dim g As Integer, r As Long
    Dim playerName As String, akaField As String
    Dim akas() As String, k As Integer

    For g = 0 To 3
        For r = 2 To 200
            playerName = Trim(SafeStr(wsRoster.Cells(r, nameCols(g)).Value))
            If playerName = "" Then GoTo NextRow1
            If NormalizeName(playerName) = normRaw Then
                FindRosterName = playerName
                Exit Function
            End If
            akaField = Trim(SafeStr(wsRoster.Cells(r, akaCols(g)).Value))
            If akaField <> "" Then
                akas = Split(akaField, ",")
                For k = 0 To UBound(akas)
                    If NormalizeName(Trim(akas(k))) = normRaw Then
                        FindRosterName = playerName
                        Exit Function
                    End If
                Next k
            End If
NextRow1:
        Next r
    Next g

    Dim bestName As String
    Dim bestScore As Double
    bestScore = 0
    Dim sim As Double

    For g = 0 To 3
        For r = 2 To 200
            playerName = Trim(SafeStr(wsRoster.Cells(r, nameCols(g)).Value))
            If playerName = "" Then GoTo NextRow2
            sim = NameSimilarity(normRaw, NormalizeName(playerName))
            If sim > bestScore Then bestScore = sim: bestName = playerName
            akaField = Trim(SafeStr(wsRoster.Cells(r, akaCols(g)).Value))
            If akaField <> "" Then
                akas = Split(akaField, ",")
                For k = 0 To UBound(akas)
                    sim = NameSimilarity(normRaw, NormalizeName(Trim(akas(k))))
                    If sim > bestScore Then bestScore = sim: bestName = playerName
                Next k
            End If
NextRow2:
        Next r
    Next g

    If bestScore >= 0.7 Then
        FindRosterName = bestName
    Else
        FindRosterName = ""
    End If
End Function

'=======================================================
' NormalizeName
'=======================================================
Private Function NormalizeName(s As String) As String
    Dim result As String
    result = ""
    Dim i As Long
    For i = 1 To Len(s)
        Dim c As String
        c = Mid(s, i, 1)
        Dim code As Long
        code = AscW(c)
        Select Case code
            Case 65 To 90:   result = result & LCase(c)
            Case 97 To 122:  result = result & c
            Case 48:  result = result & "o"
            Case 49:  result = result & "i"
            Case 50:  result = result & "z"
            Case 51:  result = result & "e"
            Case 52:  result = result & "a"
            Case 53:  result = result & "s"
            Case 54:  result = result & "g"
            Case 55:  result = result & "t"
            Case 56:  result = result & "b"
            Case 57:  result = result & "g"
            Case 65313 To 65338: result = result & Chr(code - 65313 + 97)
            Case 65345 To 65370: result = result & Chr(code - 65345 + 97)
            Case 913: result = result & "a"
            Case 914: result = result & "b"
            Case 915: result = result & "g"
            Case 916: result = result & "d"
            Case 917: result = result & "e"
            Case 919: result = result & "n"
            Case 920: result = result & "o"
            Case 921: result = result & "i"
            Case 922: result = result & "k"
            Case 923: result = result & "l"
            Case 924: result = result & "m"
            Case 925: result = result & "n"
            Case 927: result = result & "o"
            Case 928: result = result & "p"
            Case 929: result = result & "r"
            Case 931: result = result & "s"
            Case 932: result = result & "t"
            Case 933: result = result & "u"
            Case 935: result = result & "x"
            Case 937: result = result & "w"
            Case 945: result = result & "a"
            Case 946: result = result & "b"
            Case 949: result = result & "e"
            Case 951: result = result & "n"
            Case 953: result = result & "i"
            Case 954: result = result & "k"
            Case 955: result = result & "l"
            Case 956: result = result & "m"
            Case 957: result = result & "v"
            Case 959: result = result & "o"
            Case 961: result = result & "r"
            Case 963: result = result & "s"
            Case 964: result = result & "t"
            Case 965: result = result & "u"
            Case 969: result = result & "w"
            Case 1072: result = result & "a"
            Case 1077: result = result & "e"
            Case 1086: result = result & "o"
            Case 1088: result = result & "r"
            Case 1089: result = result & "s"
            Case 1093: result = result & "x"
            Case 1091: result = result & "y"
            Case 1040: result = result & "a"
            Case 1045: result = result & "e"
            Case 1054: result = result & "o"
            Case 1056: result = result & "r"
            Case 1057: result = result & "s"
            Case 1061: result = result & "x"
            Case 1059: result = result & "u"
            Case 1042: result = result & "b"
            Case 1052: result = result & "m"
            Case 1058: result = result & "t"
            Case 1053: result = result & "n"
            Case 1048: result = result & "i"
            Case 20043: result = result & "b"
            Case 12410: result = result & "l"
            Case 21452: result = result & "h"
            Case 12310: result = result & "r"
            Case 224, 225, 226, 227, 228, 229: result = result & "a"
            Case 192, 193, 194, 195, 196, 197: result = result & "a"
            Case 232, 233, 234, 235:      result = result & "e"
            Case 200, 201, 202, 203:      result = result & "e"
            Case 236, 237, 238, 239:      result = result & "i"
            Case 204, 205, 206, 207:      result = result & "i"
            Case 242, 243, 244, 245, 246: result = result & "o"
            Case 210, 211, 212, 213, 214: result = result & "o"
            Case 249, 250, 251, 252:      result = result & "u"
            Case 217, 218, 219, 220:      result = result & "u"
            Case 253, 255:                result = result & "y"
            Case 241, 209:                result = result & "n"
            Case 231, 199:                result = result & "c"
            Case Else:
        End Select
    Next i
    NormalizeName = result
End Function

'=======================================================
' NameSimilarity
'=======================================================
Private Function NameSimilarity(a As String, b As String) As Double
    If a = "" Or b = "" Then NameSimilarity = 0: Exit Function
    If a = b Then NameSimilarity = 1: Exit Function
    Dim maxLen As Long
    maxLen = IIf(Len(a) > Len(b), Len(a), Len(b))
    NameSimilarity = 1 - (LevenshteinDist(a, b) / maxLen)
End Function

'=======================================================
' LevenshteinDist
'=======================================================
Private Function LevenshteinDist(s As String, t As String) As Long
    Dim m As Long, n As Long
    m = Len(s): n = Len(t)
    If m = 0 Then LevenshteinDist = n: Exit Function
    If n = 0 Then LevenshteinDist = m: Exit Function
    Dim d() As Long
    ReDim d(0 To m, 0 To n)
    Dim i As Long, j As Long
    For i = 0 To m: d(i, 0) = i: Next i
    For j = 0 To n: d(0, j) = j: Next j
    For j = 1 To n
        For i = 1 To m
            If Mid(s, i, 1) = Mid(t, j, 1) Then
                d(i, j) = d(i - 1, j - 1)
            Else
                Dim a As Long, b As Long, cc As Long
                a = d(i - 1, j) + 1
                b = d(i, j - 1) + 1
                cc = d(i - 1, j - 1) + 1
                d(i, j) = IIf(a < b, IIf(a < cc, a, cc), IIf(b < cc, b, cc))
            End If
        Next i
    Next j
    LevenshteinDist = d(m, n)
End Function

'=======================================================
' SafeStr - safely convert any cell value to String
' Handles error values (#N/A, #REF!, etc.), Null, Empty
'=======================================================
Private Function SafeStr(v As Variant) As String
    If IsError(v) Then
        SafeStr = ""
    ElseIf IsNull(v) Then
        SafeStr = ""
    ElseIf IsEmpty(v) Then
        SafeStr = ""
    Else
        SafeStr = CStr(v)
    End If
End Function

'=======================================================
' SafeFormat - safely format a cell value, returns "-"
' if the value is an error or non-numeric
'=======================================================
Private Function SafeFormat(v As Variant, fmt As String) As String
    If IsError(v) Then
        SafeFormat = "-"
    ElseIf IsNull(v) Or IsEmpty(v) Then
        SafeFormat = "-"
    ElseIf IsNumeric(v) Then
        SafeFormat = Format$(CDbl(v), fmt)
    Else
        SafeFormat = CStr(v)
    End If
End Function

'=======================================================
' SetupScoreImport
'=======================================================
Sub SetupScoreImport()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet("Score Import")
    With ws
        .Cells(1, 1).Value = "Player"
        .Cells(1, 2).Value = "Score"
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(68, 114, 196)
        .Rows(1).Font.Color = RGB(255, 255, 255)
        .Columns(1).ColumnWidth = 25
        .Columns(2).ColumnWidth = 15
        .Columns(2).NumberFormat = "#,##0"
    End With
    MsgBox "Score Import sheet ready! Paste WPX clipboard into cell A2 each day.", vbInformation
End Sub
