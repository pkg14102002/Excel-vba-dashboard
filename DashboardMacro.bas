' =============================================================
'  Self-Refreshing Excel Dashboard with VBA Macro Engine
'  Author  : Prince Kumar Gupta
'  Role    : Data Analyst
'  Tools   : Advanced Excel, VBA, Power Query, SQL Connection
' =============================================================

Option Explicit

' ── Constants ────────────────────────────────────────────────────
Private Const NAVY_COLOR    As Long = 874295     ' #0D2137
Private Const BLUE_COLOR    As Long = 1398476    ' #1565C0
Private Const GREEN_COLOR   As Long = 1810720    ' #1B5E20
Private Const RED_COLOR     As Long = 12206086   ' #B71C1C
Private Const HEADER_COLOR  As Long = 15132390   ' #E3F2FD

' ── MAIN ENTRY POINT ─────────────────────────────────────────────
Public Sub RunDashboard()
    Application.ScreenUpdating = False
    Application.Calculation    = xlCalculationManual
    Application.EnableEvents   = False

    On Error GoTo ErrorHandler

    Call ShowProgress("Starting Dashboard Engine...", 0)
    Call ClearDashboard
    Call ShowProgress("Generating sample data...", 15)
    Call GenerateSampleData
    Call ShowProgress("Building KPI section...", 35)
    Call BuildKPISection
    Call ShowProgress("Building regional summary...", 55)
    Call BuildRegionalSummary
    Call ShowProgress("Building product table...", 70)
    Call BuildProductTable
    Call ShowProgress("Applying conditional formatting...", 85)
    Call ApplyConditionalFormatting
    Call ShowProgress("Adding charts...", 90)
    Call AddRevenueChart
    Call ShowProgress("Finalising dashboard...", 98)
    Call FinaliseLayout
    Call ShowProgress("Done!", 100)

    MsgBox "Dashboard refreshed successfully!" & Chr(13) & _
           "Last Updated: " & Now(), vbInformation, "Dashboard Engine"
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
Finally:
    Application.ScreenUpdating = True
    Application.Calculation    = xlCalculationAutomatic
    Application.EnableEvents   = True
End Sub

' ── Generate Sample Data ─────────────────────────────────────────
Private Sub GenerateSampleData()
    Dim wsData As Worksheet
    Dim regions()  As String: regions  = Split("North,South,East,West,Central", ",")
    Dim products() As String: products = Split("Product A,Product B,Product C,Product D", ",")
    Dim statuses() As String: statuses = Split("Closed,Open,Pending", ",")

    On Error Resume Next
    Set wsData = ThisWorkbook.Sheets("RawData")
    If wsData Is Nothing Then
        Set wsData = ThisWorkbook.Sheets.Add
        wsData.Name = "RawData"
    End If
    On Error GoTo 0

    wsData.Cells.Clear
    wsData.Visible = xlSheetVeryHidden  ' Hide from users

    ' Headers
    Dim headers() As String
    headers = Split("Date,Region,Product,Revenue,Units,Target,Status,Achievement%", ",")
    Dim i As Integer
    For i = 0 To UBound(headers)
        With wsData.Cells(1, i + 1)
            .Value          = headers(i)
            .Font.Bold      = True
            .Interior.Color = NAVY_COLOR
            .Font.Color     = RGB(255, 255, 255)
        End With
    Next i

    ' Data rows
    Randomize
    Dim r As Long, revenue As Double, target As Double
    For r = 2 To 501
        revenue = Int(Rnd() * 490000) + 10000
        target  = revenue * (0.8 + Rnd() * 0.5)

        wsData.Cells(r, 1) = Date - Int(Rnd() * 90)        ' Date
        wsData.Cells(r, 2) = regions(Int(Rnd() * 5))       ' Region
        wsData.Cells(r, 3) = products(Int(Rnd() * 4))      ' Product
        wsData.Cells(r, 4) = Round(revenue, 0)              ' Revenue
        wsData.Cells(r, 5) = Int(Rnd() * 99) + 1           ' Units
        wsData.Cells(r, 6) = Round(target, 0)               ' Target
        wsData.Cells(r, 7) = statuses(Int(Rnd() * 3))      ' Status
        wsData.Cells(r, 8) = Round(revenue / target * 100, 1) ' Achievement %
    Next r
End Sub

' ── Clear Dashboard ──────────────────────────────────────────────
Private Sub ClearDashboard()
    Dim wsDash As Worksheet
    Set wsDash = GetOrCreateSheet("Dashboard")
    wsDash.Cells.Clear
    wsDash.Cells.Interior.Color = RGB(248, 250, 252)
    wsDash.Activate
End Sub

' ── Build KPI Section ────────────────────────────────────────────
Private Sub BuildKPISection()
    Dim wsDash As Worksheet, wsData As Worksheet
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set wsData = ThisWorkbook.Sheets("RawData")

    ' Title Banner
    With wsDash.Range("A1:L2")
        .Merge
        .Value                  = "SALES PERFORMANCE DASHBOARD"
        .Font.Bold              = True
        .Font.Size              = 18
        .Font.Color             = RGB(255, 255, 255)
        .Interior.Color         = NAVY_COLOR
        .HorizontalAlignment    = xlCenter
        .VerticalAlignment      = xlCenter
    End With
    wsDash.Rows("1:2").RowHeight = 22

    ' Subtitle
    With wsDash.Range("A3:L3")
        .Merge
        .Value               = "Last Updated: " & Now() & "  |  Prepared by: Prince Kumar Gupta | Data Analyst"
        .Font.Size           = 9
        .Font.Italic         = True
        .Font.Color          = RGB(96, 125, 139)
        .HorizontalAlignment = xlCenter
        .Interior.Color      = RGB(236, 242, 255)
    End With

    ' KPI Cards
    Dim totalRev  As Double: totalRev  = Application.WorksheetFunction.Sum(wsData.Range("D2:D501"))
    Dim totalUnits As Long:  totalUnits = Application.WorksheetFunction.Sum(wsData.Range("E2:E501"))
    Dim avgAch    As Double: avgAch    = Application.WorksheetFunction.Average(wsData.Range("H2:H501"))
    Dim closedDeals As Long
    Dim r As Long
    For r = 2 To 501
        If wsData.Cells(r, 7).Value = "Closed" Then closedDeals = closedDeals + 1
    Next r

    Dim kpiTitles()  As String
    Dim kpiValues()  As String
    Dim kpiColors()  As Long

    kpiTitles = Split("Total Revenue|Total Units Sold|Avg Achievement|Closed Deals", "|")

    ReDim kpiValues(3)
    kpiValues(0) = "Rs." & Format(totalRev, "#,##0")
    kpiValues(1) = Format(totalUnits, "#,##0")
    kpiValues(2) = Format(avgAch, "0.0") & "%"
    kpiValues(3) = Format(closedDeals, "#,##0")

    kpiColors = Array(RGB(21, 101, 192), RGB(27, 94, 32), RGB(230, 81, 0), RGB(74, 20, 140))

    Dim col As Integer
    For col = 0 To 3
        Dim startCol As Integer: startCol = col * 3 + 1
        With wsDash.Range(wsDash.Cells(5, startCol), wsDash.Cells(7, startCol + 1))
            .Merge
            .Interior.Color         = kpiColors(col)
            .Font.Color             = RGB(255, 255, 255)
            .HorizontalAlignment    = xlCenter
            .VerticalAlignment      = xlCenter
        End With
        wsDash.Cells(5, startCol).Value = kpiTitles(col)
        wsDash.Cells(5, startCol).Font.Size = 9
        wsDash.Cells(6, startCol).Value = kpiValues(col)
        wsDash.Cells(6, startCol).Font.Size = 14
        wsDash.Cells(6, startCol).Font.Bold = True
    Next col
    wsDash.Rows("5:7").RowHeight = 18
End Sub

' ── Build Regional Summary Table ─────────────────────────────────
Private Sub BuildRegionalSummary()
    Dim wsDash As Worksheet, wsData As Worksheet
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set wsData = ThisWorkbook.Sheets("RawData")

    wsDash.Cells(9, 1).Value = "REGIONAL PERFORMANCE"
    With wsDash.Cells(9, 1)
        .Font.Bold  = True
        .Font.Size  = 11
        .Font.Color = NAVY_COLOR
    End With

    Dim headers() As String
    headers = Split("Region,Revenue (Rs.),Units,Achievement %,Deals", ",")
    Dim c As Integer
    For c = 0 To UBound(headers)
        With wsDash.Cells(10, c + 1)
            .Value          = headers(c)
            .Font.Bold      = True
            .Font.Color     = RGB(255, 255, 255)
            .Interior.Color = BLUE_COLOR
            .HorizontalAlignment = xlCenter
        End With
    Next c

    Dim regions() As String: regions = Split("North,South,East,West,Central", ",")
    Dim row As Integer: row = 11
    Dim reg As Variant

    For Each reg In regions
        Dim revSum As Double: revSum  = 0
        Dim unitSum As Long:  unitSum = 0
        Dim achSum  As Double: achSum = 0
        Dim cnt     As Long:  cnt    = 0
        Dim i As Long

        For i = 2 To 501
            If wsData.Cells(i, 2).Value = reg Then
                revSum  = revSum  + wsData.Cells(i, 4).Value
                unitSum = unitSum + wsData.Cells(i, 5).Value
                achSum  = achSum  + wsData.Cells(i, 8).Value
                cnt = cnt + 1
            End If
        Next i

        wsDash.Cells(row, 1) = reg
        wsDash.Cells(row, 2) = Format(revSum, "#,##0")
        wsDash.Cells(row, 3) = Format(unitSum, "#,##0")
        wsDash.Cells(row, 4) = Format(IIf(cnt > 0, achSum / cnt, 0), "0.0") & "%"
        wsDash.Cells(row, 5) = cnt

        If row Mod 2 = 0 Then
            wsDash.Range(wsDash.Cells(row, 1), wsDash.Cells(row, 5)).Interior.Color = HEADER_COLOR
        End If
        row = row + 1
    Next reg
End Sub

' ── Build Product Table ──────────────────────────────────────────
Private Sub BuildProductTable()
    Dim wsDash As Worksheet, wsData As Worksheet
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set wsData = ThisWorkbook.Sheets("RawData")

    wsDash.Cells(9, 7).Value = "PRODUCT PERFORMANCE"
    With wsDash.Cells(9, 7)
        .Font.Bold  = True
        .Font.Size  = 11
        .Font.Color = NAVY_COLOR
    End With

    Dim products() As String: products = Split("Product A,Product B,Product C,Product D", ",")
    Dim headers()  As String: headers  = Split("Product,Revenue,Units,Avg Deal", ",")
    Dim c As Integer
    For c = 0 To UBound(headers)
        With wsDash.Cells(10, c + 7)
            .Value          = headers(c)
            .Font.Bold      = True
            .Font.Color     = RGB(255, 255, 255)
            .Interior.Color = GREEN_COLOR
            .HorizontalAlignment = xlCenter
        End With
    Next c

    Dim row As Integer: row = 11
    Dim prod As Variant
    For Each prod In products
        Dim revS As Double: revS  = 0
        Dim untS As Long:   untS  = 0
        Dim cnt  As Long:   cnt   = 0
        Dim i As Long
        For i = 2 To 501
            If wsData.Cells(i, 3).Value = prod Then
                revS = revS + wsData.Cells(i, 4).Value
                untS = untS + wsData.Cells(i, 5).Value
                cnt  = cnt  + 1
            End If
        Next i
        wsDash.Cells(row, 7)  = prod
        wsDash.Cells(row, 8)  = Format(revS, "#,##0")
        wsDash.Cells(row, 9)  = Format(untS, "#,##0")
        wsDash.Cells(row, 10) = Format(IIf(cnt > 0, revS / cnt, 0), "#,##0")
        If row Mod 2 = 0 Then
            wsDash.Range(wsDash.Cells(row, 7), wsDash.Cells(row, 10)).Interior.Color = RGB(232, 245, 233)
        End If
        row = row + 1
    Next prod
End Sub

' ── Conditional Formatting ───────────────────────────────────────
Private Sub ApplyConditionalFormatting()
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Dim rng As Range
    Set rng = wsDash.Range("D11:D15") ' Achievement column

    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(xlCellValue, xlGreaterEqual, 100)
        .Interior.Color = RGB(200, 230, 201)
        .Font.Color     = GREEN_COLOR
        .Font.Bold      = True
    End With
    With rng.FormatConditions.Add(xlCellValue, xlLess, 75)
        .Interior.Color = RGB(255, 235, 238)
        .Font.Color     = RED_COLOR
        .Font.Bold      = True
    End With
End Sub

' ── Add Revenue Chart ────────────────────────────────────────────
Private Sub AddRevenueChart()
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets("Dashboard")

    Dim cht As ChartObject
    Set cht = wsDash.ChartObjects.Add(Left:=10, Top:=280, Width:=380, Height:=200)
    With cht.Chart
        .ChartType = xlColumnClustered
        .SetSourceData wsDash.Range("A10:B15")
        .HasTitle        = True
        .ChartTitle.Text = "Revenue by Region"
        .ChartTitle.Font.Size  = 11
        .ChartTitle.Font.Bold  = True
        .ChartTitle.Font.Color = NAVY_COLOR
        .PlotArea.Interior.Color = RGB(248, 250, 252)
        .ChartArea.Border.LineStyle = xlNone
    End With
End Sub

' ── Finalise Layout ──────────────────────────────────────────────
Private Sub FinaliseLayout()
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets("Dashboard")

    With wsDash
        .Columns("A:L").AutoFit
        .Rows("1:40").RowHeight = 18
        .Rows("1:2").RowHeight  = 24
        .Tab.Color              = NAVY_COLOR
        .Activate
    End With

    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings  = False
End Sub

' ── Helper: Get or Create Sheet ──────────────────────────────────
Private Function GetOrCreateSheet(name As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = name
    End If
    Set GetOrCreateSheet = ws
End Function

' ── Progress Bar ─────────────────────────────────────────────────
Private Sub ShowProgress(msg As String, pct As Integer)
    Application.StatusBar = "Dashboard Engine: " & msg & " [" & pct & "%]"
    DoEvents
End Sub
