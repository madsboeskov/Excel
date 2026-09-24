Attribute VB_Name = "CapIQ_DCF"
Option Explicit

'==============================================================================
' Capital IQ -> DCF model builder
'
' Reads historical financials that were exported / pasted from S&P Capital IQ
' into one or more worksheets and builds a fully formula-driven, unlevered
' free-cash-flow DCF on a new sheet called "DCF".
'
' How to use
'   1. Import this file into the workbook (Alt+F11 > File > Import File...).
'   2. Paste / export your Capital IQ financials (Income Statement, Balance
'      Sheet, Cash Flow - one sheet or several) into the workbook.
'   3. Run the macro BuildDCF (Alt+F8) and enter the sheet name(s) holding
'      the data, or * to search every sheet.
'
' Everything on the generated sheet is a live formula: the historical columns
' link back to the Capital IQ cells, so refreshing the Capital IQ data (same
' layout) updates the model. Yellow cells are assumptions you can overwrite.
'==============================================================================

' ---- configuration -----------------------------------------------------------
Private Const DEFAULT_DATA_SHEET As String = "CapIQ Data"
Private Const OUTPUT_SHEET As String = "DCF"
Private Const PROJ_YEARS As Long = 5          ' number of projected years (>= 2)
Private Const MAX_HIST_YEARS As Long = 5      ' most recent historical years used
Private Const HEADER_SCAN_ROWS As Long = 60   ' rows searched for the year header
Private Const MAX_SCAN_ROWS As Long = 5000
Private Const MAX_SCAN_COLS As Long = 250

' default assumptions (all editable on the generated sheet)
Private Const DEF_RISK_FREE As Double = 0.0425
Private Const DEF_ERP As Double = 0.055
Private Const DEF_BETA As Double = 1
Private Const DEF_COST_OF_DEBT As Double = 0.06
Private Const DEF_TERMINAL_GROWTH As Double = 0.025
Private Const DEF_EXIT_MULTIPLE As Double = 10
Private Const SENS_WACC_STEP As Double = 0.005
Private Const SENS_G_STEP As Double = 0.0025
Private Const SENS_MULT_STEP As Double = 1

' ---- line items ----------------------------------------------------------------
Private Const IT_REVENUE As Long = 1
Private Const IT_EBIT As Long = 2
Private Const IT_DA As Long = 3
Private Const IT_EBT As Long = 4
Private Const IT_TAX As Long = 5
Private Const IT_CAPEX As Long = 6
Private Const IT_NWC As Long = 7
Private Const IT_CASH As Long = 8
Private Const IT_DEBT As Long = 9
Private Const IT_MINORITY As Long = 10
Private Const IT_PREFERRED As Long = 11
Private Const IT_SHARES As Long = 12
Private Const IT_BETA As Long = 13
Private Const IT_PRICE As Long = 14
Private Const IT_COUNT As Long = 14

' ---- cell styles ---------------------------------------------------------------
Private Const ST_CALC As Long = 0      ' black: formula
Private Const ST_INPUT As Long = 1     ' blue on yellow: hard-coded assumption
Private Const ST_DRIVER As Long = 2    ' black on yellow: assumption pre-filled by formula
Private Const ST_LINK As Long = 3      ' green: link to the Capital IQ sheet

Private Const FMT_NUM As String = "#,##0.0_);(#,##0.0);""-""_)"
Private Const FMT_PCT As String = "0.0%_);(0.0%);""-""_)"
Private Const FMT_MULT As String = "0.0""x"""
Private Const FMT_PX As String = "#,##0.00_);(#,##0.00)"
Private Const FMT_BETA As String = "0.00"

' ---- state -------------------------------------------------------------------------
Private mWb As Workbook
Private mNSrc As Long
Private mSrc() As Worksheet
Private mData() As Variant        ' cached cell values per source sheet
Private mHdr() As Long            ' year-header row per sheet (0 = none)
Private mLbl() As Long            ' label column per sheet
Private mLastRow() As Long
Private mLastCol() As Long

Private mNHist As Long
Private mHistYear() As Long

Private mItemSheet(1 To IT_COUNT) As Long
Private mItemRow(1 To IT_COUNT) As Long
Private mItemCol(1 To IT_COUNT) As Long    ' value column for point-in-time items
Private mItemLabel(1 To IT_COUNT) As String

'==============================================================================
' Public entry points
'==============================================================================

' Interactive macro: Alt+F8 > BuildDCF
Public Sub BuildDCF()
    Dim srcSheets As String, msg As String

    If ActiveWorkbook Is Nothing Then
        MsgBox "Open the workbook that holds your Capital IQ data first.", vbExclamation, "Build DCF"
        Exit Sub
    End If

    If SheetExists(ActiveWorkbook, DEFAULT_DATA_SHEET) Then srcSheets = DEFAULT_DATA_SHEET Else srcSheets = "*"
    srcSheets = InputBox("Sheet(s) that hold the Capital IQ financials." & vbCrLf & _
                      "Separate several sheets with commas, or enter * to search every sheet.", _
                      "Build DCF", srcSheets)
    If Len(Trim$(srcSheets)) = 0 Then Exit Sub

    If SheetExists(ActiveWorkbook, OUTPUT_SHEET) Then
        If MsgBox("The sheet '" & OUTPUT_SHEET & "' already exists and will be replaced. Continue?", _
                  vbQuestion + vbOKCancel, "Build DCF") <> vbOK Then Exit Sub
    End If

    On Error GoTo Fail
    Application.ScreenUpdating = False
    msg = BuildDCFFromSheets(ActiveWorkbook, srcSheets)
    Application.ScreenUpdating = True
    MsgBox msg, vbInformation, "Build DCF"
    Exit Sub

Fail:
    Application.ScreenUpdating = True
    MsgBox "Could not build the DCF:" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Build DCF"
End Sub

' Non-interactive version (no dialogs), e.g. for calling from other macros.
' sheetList: comma-separated sheet names, or "*" for every sheet.
' Returns a summary of what was mapped. Replaces an existing DCF sheet.
Public Function BuildDCFFromSheets(ByVal wb As Workbook, ByVal sheetList As String) As String
    LoadSources wb, sheetList
    MapLineItems
    WriteModel
    BuildDCFFromSheets = MappingSummary()
End Function

'==============================================================================
' Reading the Capital IQ data
'==============================================================================

Private Sub LoadSources(ByVal wb As Workbook, ByVal sheetList As String)
    Dim ws As Worksheet, names As Variant, nm As String, i As Long

    Set mWb = wb
    mNSrc = 0
    ReDim mSrc(1 To wb.Worksheets.Count)
    ReDim mData(1 To wb.Worksheets.Count)
    ReDim mHdr(1 To wb.Worksheets.Count)
    ReDim mLbl(1 To wb.Worksheets.Count)
    ReDim mLastRow(1 To wb.Worksheets.Count)
    ReDim mLastCol(1 To wb.Worksheets.Count)

    If Trim$(sheetList) = "*" Then
        For Each ws In wb.Worksheets
            If ws.Name <> OUTPUT_SHEET Then AddSource ws
        Next ws
    Else
        names = Split(sheetList, ",")
        For i = LBound(names) To UBound(names)
            nm = Trim$(names(i))
            If Len(nm) > 0 Then
                If Not SheetExists(wb, nm) Then RaiseErr "Sheet '" & nm & "' was not found."
                If nm = OUTPUT_SHEET Then RaiseErr "The source sheet cannot be the output sheet '" & OUTPUT_SHEET & "'."
                AddSource wb.Worksheets(nm)
            End If
        Next i
    End If
    If mNSrc = 0 Then RaiseErr "No source sheets to read."
End Sub

Private Sub AddSource(ByVal ws As Worksheet)
    Dim lastR As Long, lastC As Long

    With ws.UsedRange
        lastR = .Row + .Rows.Count - 1
        lastC = .Column + .Columns.Count - 1
    End With
    If lastR > MAX_SCAN_ROWS Then lastR = MAX_SCAN_ROWS
    If lastC > MAX_SCAN_COLS Then lastC = MAX_SCAN_COLS
    If lastR < 2 Then lastR = 2
    If lastC < 2 Then lastC = 2

    mNSrc = mNSrc + 1
    Set mSrc(mNSrc) = ws
    mData(mNSrc) = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value
    mLastRow(mNSrc) = lastR
    mLastCol(mNSrc) = lastC
    DetectLayout mData(mNSrc), mNSrc
End Sub

' Finds the row holding the fiscal-year headers and the column holding the
' line-item labels.
Private Sub DetectLayout(ByRef d As Variant, ByVal s As Long)
    Dim r As Long, c As Long, n As Long, best As Long, maxR As Long, lastLblCol As Long

    mHdr(s) = 0
    best = 1                      ' need at least two year columns
    maxR = mLastRow(s)
    If maxR > HEADER_SCAN_ROWS Then maxR = HEADER_SCAN_ROWS
    For r = 1 To maxR
        n = 0
        For c = 2 To mLastCol(s)
            If HeaderYear(d(r, c)) > 0 Then n = n + 1
        Next c
        If n > best Then best = n: mHdr(s) = r
    Next r

    lastLblCol = mLastCol(s)
    If mHdr(s) > 0 Then
        For c = 2 To mLastCol(s)
            If HeaderYear(d(mHdr(s), c)) > 0 Then lastLblCol = c - 1: Exit For
        Next c
    End If
    If lastLblCol > 10 Then lastLblCol = 10

    mLbl(s) = 1
    best = -1
    For c = 1 To lastLblCol
        n = 0
        For r = mHdr(s) + 1 To mLastRow(s)
            If IsTextCell(d(r, c)) Then n = n + 1
        Next r
        If n > best Then best = n: mLbl(s) = c
    Next c
End Sub

Private Sub MapLineItems()
    Dim i As Long

    For i = 1 To IT_COUNT
        mItemSheet(i) = 0: mItemRow(i) = 0: mItemCol(i) = 0: mItemLabel(i) = ""
    Next i

    FindItem IT_REVENUE
    If mItemRow(IT_REVENUE) = 0 Then
        RaiseErr "Could not find a revenue line (e.g. 'Total Revenue') with values under a fiscal-year header " & _
                 "(e.g. 'FY2023', 'Dec-31-2023' or '12 months Dec-31-2023')."
    End If
    CollectHistYears

    For i = 1 To IT_COUNT
        If i <> IT_REVENUE Then FindItem i
    Next i
    If mItemRow(IT_EBIT) = 0 Then RaiseErr "Could not find an EBIT / 'Operating Income' line."

    For i = 1 To IT_COUNT
        If mItemRow(i) > 0 And Not IsPeriodItem(i) Then mItemCol(i) = PointValueCol(i)
    Next i
End Sub

' Historical years = years in which the revenue line has a value.
Private Sub CollectHistYears()
    Dim s As Long, r As Long, c As Long, y As Long, i As Long, j As Long, n As Long, tmp As Long
    Dim yrs() As Long

    s = mItemSheet(IT_REVENUE)
    r = mItemRow(IT_REVENUE)
    ReDim yrs(1 To mLastCol(s))
    n = 0
    For c = mLbl(s) + 1 To mLastCol(s)
        y = HeaderYear(mData(s)(mHdr(s), c))
        If y > 0 And Not IsEstimateHeader(mData(s)(mHdr(s), c)) Then
            If IsNum(mData(s)(r, c)) Then
                For i = 1 To n
                    If yrs(i) = y Then Exit For
                Next i
                If i > n Then n = n + 1: yrs(n) = y
            End If
        End If
    Next c
    If n = 0 Then RaiseErr "The revenue line has no values in historical fiscal-year columns."

    For i = 2 To n                       ' insertion sort, ascending
        tmp = yrs(i)
        j = i - 1
        Do While j >= 1
            If yrs(j) <= tmp Then Exit Do
            yrs(j + 1) = yrs(j)
            j = j - 1
        Loop
        yrs(j + 1) = tmp
    Next i

    mNHist = n
    If mNHist > MAX_HIST_YEARS Then mNHist = MAX_HIST_YEARS
    ReDim mHistYear(1 To mNHist)
    For i = 1 To mNHist
        mHistYear(i) = yrs(n - mNHist + i)
    Next i
End Sub

' Searches all source sheets for a line item. Pass 1 = exact label match,
' pass 2 = label starts with the synonym, pass 3 = label contains it.
' Synonym order decides priority within a pass.
Private Sub FindItem(ByVal item As Long)
    Dim syns As Variant, pass As Long, k As Long, s As Long, r As Long, ns As String

    syns = ItemSynonyms(item)
    For pass = 1 To 3
        For k = LBound(syns) To UBound(syns)
            ns = NormLabel(CStr(syns(k)))
            If pass = 1 Or InStr(ns, " ") > 0 Then     ' fuzzy passes only for multi-word labels
                For s = 1 To mNSrc
                    r = FindRowInSheet(mData(s), s, item, ns, pass)
                    If r > 0 Then
                        mItemSheet(item) = s
                        mItemRow(item) = r
                        mItemLabel(item) = Trim$(CStr(mData(s)(r, mLbl(s))))
                        Exit Sub
                    End If
                Next s
            End If
        Next k
    Next pass
End Sub

Private Function FindRowInSheet(ByRef d As Variant, ByVal s As Long, ByVal item As Long, _
                                ByVal ns As String, ByVal pass As Long) As Long
    Dim r As Long, raw As String, nl As String, ok As Boolean

    For r = 1 To mLastRow(s)
        If r <> mHdr(s) Then
            If IsTextCell(d(r, mLbl(s))) Then
                raw = CStr(d(r, mLbl(s)))
                nl = NormLabel(raw)
                Select Case pass
                    Case 1: ok = (nl = ns)
                    Case 2: ok = (Left$(nl, Len(ns) + 1) = ns & " ")
                    Case Else: ok = (InStr(" " & nl & " ", " " & ns & " ") > 0)
                End Select
                If ok And pass > 1 Then ok = Not IsExcludedLabel(raw)
                If ok Then ok = RowHasData(d, s, r, item)
                If ok Then FindRowInSheet = r: Exit Function
            End If
        End If
    Next r
End Function

Private Function RowHasData(ByRef d As Variant, ByVal s As Long, ByVal r As Long, ByVal item As Long) As Boolean
    Dim c As Long, i As Long

    If IsPeriodItem(item) Then
        If mHdr(s) = 0 Then Exit Function
        If item = IT_REVENUE Then
            For c = mLbl(s) + 1 To mLastCol(s)
                If HeaderYear(d(mHdr(s), c)) > 0 And Not IsEstimateHeader(d(mHdr(s), c)) Then
                    If IsNum(d(r, c)) Then RowHasData = True: Exit Function
                End If
            Next c
        Else
            For i = 1 To mNHist
                c = YearCol(d, s, mHistYear(i))
                If c > 0 Then
                    If IsNum(d(r, c)) Then RowHasData = True: Exit Function
                End If
            Next i
        End If
    Else
        For c = mLbl(s) + 1 To mLastCol(s)
            If IsNum(d(r, c)) Then RowHasData = True: Exit Function
        Next c
    End If
End Function

' Column of a historical fiscal year on a sheet (0 if absent).
Private Function YearCol(ByRef d As Variant, ByVal s As Long, ByVal y As Long) As Long
    Dim c As Long
    If mHdr(s) = 0 Then Exit Function
    For c = mLbl(s) + 1 To mLastCol(s)
        If HeaderYear(d(mHdr(s), c)) = y Then
            If Not IsEstimateHeader(d(mHdr(s), c)) Then YearCol = c: Exit Function
        End If
    Next c
End Function

' Point-in-time items (cash, debt, shares, ...): latest historical year if
' available, otherwise the right-most number in the row (e.g. beta, price).
Private Function PointValueCol(ByVal item As Long) As Long
    Dim s As Long, r As Long, c As Long

    s = mItemSheet(item)
    r = mItemRow(item)
    If item <> IT_BETA And item <> IT_PRICE Then
        c = YearCol(mData(s), s, mHistYear(mNHist))
        If c > 0 Then
            If IsNum(mData(s)(r, c)) Then PointValueCol = c: Exit Function
        End If
    End If
    For c = mLastCol(s) To mLbl(s) + 1 Step -1
        If IsNum(mData(s)(r, c)) Then PointValueCol = c: Exit Function
    Next c
End Function

'==============================================================================
' Line-item dictionary - add your own labels here if Capital IQ (or your
' template) uses different captions.
'==============================================================================

Private Function ItemSynonyms(ByVal item As Long) As Variant
    Select Case item
        Case IT_REVENUE
            ItemSynonyms = Array("Total Revenue", "Total Revenues", "Revenue", "Revenues", "Net Revenue", _
                                 "Net Revenues", "Net Sales", "Total Net Sales", "Sales", "Turnover")
        Case IT_EBIT
            ItemSynonyms = Array("EBIT", "Operating Income", "Operating Profit", "Total Operating Income", _
                                 "Income from Operations")
        Case IT_DA
            ItemSynonyms = Array("Depreciation & Amort., Total", "Depreciation & Amortization, Total", _
                                 "Depreciation & Amort.", "Depreciation & Amortization", _
                                 "Depreciation and Amortization", "D&A")
        Case IT_EBT
            ItemSynonyms = Array("EBT, Incl. Unusual Items", "EBT, Excl. Unusual Items", "EBT", _
                                 "Pre-Tax Income", "Pretax Income", "Income Before Taxes", _
                                 "Income Before Income Taxes", "Earnings Before Taxes")
        Case IT_TAX
            ItemSynonyms = Array("Income Tax Expense", "Income Taxes", "Provision for Income Taxes", _
                                 "Income Tax", "Tax Expense")
        Case IT_CAPEX
            ItemSynonyms = Array("Capital Expenditure", "Capital Expenditures", "Capex", _
                                 "Purchase of Property, Plant & Equipment", "Purchases of PP&E")
        Case IT_NWC
            ItemSynonyms = Array("Change in Net Working Capital", "Change in Working Capital", _
                                 "Changes in Working Capital", "(Increase) Decrease in Working Capital")
        Case IT_CASH
            ItemSynonyms = Array("Total Cash & ST Investments", "Total Cash And Short Term Investments", _
                                 "Cash And Equivalents", "Cash & Equivalents", "Cash and Cash Equivalents")
        Case IT_DEBT
            ItemSynonyms = Array("Total Debt", "Total Borrowings")
        Case IT_MINORITY
            ItemSynonyms = Array("Minority Interest", "Total Minority Interest", "Non-Controlling Interest", _
                                 "Noncontrolling Interest", "Non-Controlling Interests")
        Case IT_PREFERRED
            ItemSynonyms = Array("Total Preferred Equity", "Preferred Equity", "Preferred Stock")
        Case IT_SHARES
            ItemSynonyms = Array("Weighted Avg. Diluted Shares Out.", "Weighted Average Diluted Shares Outstanding", _
                                 "Diluted Shares Outstanding", "Total Shares Out. on Filing Date", _
                                 "Total Shares Out. on Balance Sheet Date", "Shares Outstanding")
        Case IT_BETA
            ItemSynonyms = Array("Beta", "Levered Beta", "5 Year Beta", "Beta 5 Year")
        Case IT_PRICE
            ItemSynonyms = Array("Share Price", "Stock Price", "Last Close Price", "Current Share Price", _
                                 "Last Sale Price", "Price")
    End Select
End Function

Private Function ItemName(ByVal item As Long) As String
    Select Case item
        Case IT_REVENUE: ItemName = "Revenue"
        Case IT_EBIT: ItemName = "EBIT"
        Case IT_DA: ItemName = "Depreciation & amortization"
        Case IT_EBT: ItemName = "Pre-tax income (EBT)"
        Case IT_TAX: ItemName = "Income tax expense"
        Case IT_CAPEX: ItemName = "Capital expenditure"
        Case IT_NWC: ItemName = "Change in net working capital"
        Case IT_CASH: ItemName = "Cash & ST investments"
        Case IT_DEBT: ItemName = "Total debt"
        Case IT_MINORITY: ItemName = "Minority interest"
        Case IT_PREFERRED: ItemName = "Preferred equity"
        Case IT_SHARES: ItemName = "Diluted shares outstanding"
        Case IT_BETA: ItemName = "Beta"
        Case IT_PRICE: ItemName = "Share price"
    End Select
End Function

Private Function IsPeriodItem(ByVal item As Long) As Boolean
    IsPeriodItem = (item <= IT_NWC)
End Function

'==============================================================================
' Writing the model
'==============================================================================

Private Sub WriteModel()
    Dim ws As Worksheet, P As Long, i As Long, j As Long, c As Long, r As Long, y As Long
    Dim cH1 As Long, cHN As Long, cP1 As Long, cPN As Long, cLast As Long
    Dim ebt As String, tx As String, rowNum As Variant

    ' assumption rows
    Dim rPrice As Long, rShares As Long, rRf As Long, rErp As Long, rBeta As Long, rKe As Long
    Dim rKd As Long, rTax As Long, rKdAt As Long, rMcap As Long, rDebtA As Long, rWe As Long
    Dim rWd As Long, rWacc As Long, rCagr As Long, rG As Long, rMult As Long, rMethod As Long, rMid As Long
    ' projection rows
    Dim rHdr As Long, rRev As Long, rGrowth As Long, rEbitda As Long, rEbitdaM As Long, rDA As Long
    Dim rDAPct As Long, rEbit As Long, rEbitM As Long, rTaxRate As Long, rTaxes As Long, rNopat As Long
    Dim rDA2 As Long, rCapex As Long, rCapexPct As Long, rNwc As Long, rNwcPct As Long, rFcf As Long
    Dim rT As Long, rDiscP As Long, rDF As Long, rPV As Long
    ' valuation rows
    Dim rSumPV As Long, rTvG As Long, rTvX As Long, rTv As Long, rPvTv As Long, rEV As Long, rTvPct As Long
    Dim rDebtV As Long, rCashV As Long, rMiV As Long, rPrefV As Long, rEqV As Long, rSharesV As Long
    Dim rPxV As Long, rCurPx As Long, rUpside As Long, rImpEv As Long, rImpMult As Long, rImpG As Long
    Dim rSens1 As Long, rSens2 As Long, rMap As Long

    P = PROJ_YEARS
    If P < 2 Then P = 2
    cH1 = 3
    cHN = cH1 + mNHist - 1
    cP1 = cHN + 1
    cPN = cP1 + P - 1
    cLast = cPN
    If cLast < 8 Then cLast = 8

    ' ---- row layout -------------------------------------------------------------
    r = 5
    rPrice = r: r = r + 1
    rShares = r: r = r + 1
    rRf = r: r = r + 1
    rErp = r: r = r + 1
    rBeta = r: r = r + 1
    rKe = r: r = r + 1
    rKd = r: r = r + 1
    rTax = r: r = r + 1
    rKdAt = r: r = r + 1
    rMcap = r: r = r + 1
    rDebtA = r: r = r + 1
    rWe = r: r = r + 1
    rWd = r: r = r + 1
    rWacc = r: r = r + 2
    rCagr = r: r = r + 1
    rG = r: r = r + 1
    rMult = r: r = r + 1
    rMethod = r: r = r + 1
    rMid = r: r = r + 2

    rHdr = r + 1: r = r + 2
    rRev = r: r = r + 1
    rGrowth = r: r = r + 1
    rEbitda = r: r = r + 1
    rEbitdaM = r: r = r + 1
    rDA = r: r = r + 1
    rDAPct = r: r = r + 1
    rEbit = r: r = r + 1
    rEbitM = r: r = r + 2
    rTaxRate = r: r = r + 1
    rTaxes = r: r = r + 1
    rNopat = r: r = r + 1
    rDA2 = r: r = r + 1
    rCapex = r: r = r + 1
    rCapexPct = r: r = r + 1
    rNwc = r: r = r + 1
    rNwcPct = r: r = r + 1
    rFcf = r: r = r + 2
    rT = r: r = r + 1
    rDiscP = r: r = r + 1
    rDF = r: r = r + 1
    rPV = r: r = r + 2

    r = r + 1                            ' section header
    rSumPV = r: r = r + 1
    rTvG = r: r = r + 1
    rTvX = r: r = r + 1
    rTv = r: r = r + 1
    rPvTv = r: r = r + 1
    rEV = r: r = r + 1
    rTvPct = r: r = r + 2
    rDebtV = r: r = r + 1
    rCashV = r: r = r + 1
    rMiV = r: r = r + 1
    rPrefV = r: r = r + 1
    rEqV = r: r = r + 1
    rSharesV = r: r = r + 1
    rPxV = r: r = r + 1
    rCurPx = r: r = r + 1
    rUpside = r: r = r + 2
    rImpEv = r: r = r + 1
    rImpMult = r: r = r + 1
    rImpG = r: r = r + 2
    rSens1 = r + 1: r = r + 10
    rSens2 = r: r = r + 9
    rMap = r + 1

    ' ---- sheet -----------------------------------------------------------------------
    Application.DisplayAlerts = False
    If SheetExists(mWb, OUTPUT_SHEET) Then mWb.Worksheets(OUTPUT_SHEET).Delete
    Application.DisplayAlerts = True
    Set ws = mWb.Worksheets.Add(After:=mSrc(mNSrc))
    ws.Name = OUTPUT_SHEET
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10

    ws.Cells(1, 1).Value = "Discounted Cash Flow Model"
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(2, 1).Value = "Source: " & SourceNames() & "  |  built " & Format$(Date, "yyyy-mm-dd") & _
                           "  |  figures in the units of the Capital IQ data (typically millions)"
    ws.Cells(2, 1).Font.Italic = True
    ws.Cells(3, 1).Value = "Yellow = assumption you can overwrite, blue = hard-coded input, " & _
                           "green = link to Capital IQ data, black = formula"
    ws.Cells(3, 1).Font.Italic = True
    ws.Cells(3, 1).Font.Color = RGB(89, 89, 89)

    ' ---- assumptions -----------------------------------------------------------------
    SectionHeader ws, 4, "KEY ASSUMPTIONS", cLast

    PutPoint ws, rPrice, "Current share price", IT_PRICE, 0, FMT_PX, "Enter the current share price"
    PutPoint ws, rShares, "Diluted shares outstanding", IT_SHARES, 0, FMT_NUM, _
             "Same unit as the financials (Capital IQ: millions)"
    PutAssumption ws, rRf, "Risk-free rate", DEF_RISK_FREE, FMT_PCT, ST_INPUT, "e.g. 10-year government bond yield"
    PutAssumption ws, rErp, "Equity risk premium", DEF_ERP, FMT_PCT, ST_INPUT
    PutPoint ws, rBeta, "Levered beta", IT_BETA, DEF_BETA, FMT_BETA, "Default 1.00 - no beta found in the data"
    PutAssumption ws, rKe, "Cost of equity", "=" & AbsA(rRf, 3) & "+" & AbsA(rBeta, 3) & "*" & AbsA(rErp, 3), _
                  FMT_PCT, ST_CALC, "CAPM"
    PutAssumption ws, rKd, "Pre-tax cost of debt", DEF_COST_OF_DEBT, FMT_PCT, ST_INPUT
    PutAssumption ws, rTax, "Tax rate", "=IFERROR(MIN(MAX(AVERAGE(" & RangeA(rTaxRate, cH1, cHN) & "),0),0.4),0.25)", _
                  FMT_PCT, ST_DRIVER, "Default: average historical effective tax rate (25% if n/a)"
    PutAssumption ws, rKdAt, "After-tax cost of debt", "=" & AbsA(rKd, 3) & "*(1-" & AbsA(rTax, 3) & ")", FMT_PCT, ST_CALC
    PutAssumption ws, rMcap, "Market capitalisation", "=" & AbsA(rPrice, 3) & "*" & AbsA(rShares, 3), FMT_NUM, ST_CALC
    PutPoint ws, rDebtA, "Total debt", IT_DEBT, 0, FMT_NUM, "No debt line found - enter manually"
    PutAssumption ws, rWe, "Equity / (debt + equity)", _
                  "=IF(" & AbsA(rMcap, 3) & ">0," & AbsA(rMcap, 3) & "/(" & AbsA(rMcap, 3) & "+" & AbsA(rDebtA, 3) & "),1)", FMT_PCT, ST_DRIVER, _
                  "Default: market-value weights; overwrite with a target capital structure if preferred"
    PutAssumption ws, rWd, "Debt / (debt + equity)", "=1-" & AbsA(rWe, 3), FMT_PCT, ST_CALC
    PutAssumption ws, rWacc, "WACC", "=" & AbsA(rWe, 3) & "*" & AbsA(rKe, 3) & "+" & AbsA(rWd, 3) & "*" & AbsA(rKdAt, 3), _
                  FMT_PCT, ST_CALC
    ws.Range(ws.Cells(rWacc, 1), ws.Cells(rWacc, 3)).Font.Bold = True
    TopBorder ws.Range(ws.Cells(rWacc, 1), ws.Cells(rWacc, 3))

    If mNHist >= 2 Then
        PutAssumption ws, rCagr, "Historical revenue CAGR", "=IFERROR((" & AbsA(rRev, cHN) & "/" & AbsA(rRev, cH1) & _
                      ")^(1/" & (mNHist - 1) & ")-1,0.05)", FMT_PCT, ST_CALC, _
                      "FY" & mHistYear(1) & "-FY" & mHistYear(mNHist) & "; projected growth fades from this to the terminal rate"
    Else
        PutAssumption ws, rCagr, "Starting revenue growth", 0.05, FMT_PCT, ST_INPUT, _
                      "Only one historical year - enter the expected near-term growth"
    End If
    PutAssumption ws, rG, "Terminal growth rate", DEF_TERMINAL_GROWTH, FMT_PCT, ST_INPUT, "Gordon growth method"
    PutAssumption ws, rMult, "Exit EV / EBITDA multiple", DEF_EXIT_MULTIPLE, FMT_MULT, ST_INPUT, "Exit multiple method"
    PutAssumption ws, rMethod, "Terminal value method", 1, "0", ST_INPUT, _
                  "1 = Gordon growth, 2 = exit multiple, 3 = average of both"
    PutAssumption ws, rMid, "Mid-year convention", 1, "0", ST_INPUT, "1 = yes, 0 = no (year-end discounting)"

    ' ---- projections -------------------------------------------------------------------
    SectionHeader ws, rHdr - 1, "FREE CASH FLOW PROJECTION", cLast
    ws.Cells(rHdr, 1).Value = "Fiscal year"
    For i = 1 To mNHist
        ws.Cells(rHdr, cH1 + i - 1).Value = "FY" & mHistYear(i) & "A"
    Next i
    For j = 1 To P
        ws.Cells(rHdr, cP1 + j - 1).Value = "FY" & (mHistYear(mNHist) + j) & "E"
    Next j
    With ws.Range(ws.Cells(rHdr, 1), ws.Cells(rHdr, cPN))
        .Font.Bold = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    ws.Range(ws.Cells(rHdr, cH1), ws.Cells(rHdr, cPN)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(rHdr, cP1), ws.Cells(rHdr, cPN)).Interior.Color = RGB(221, 235, 247)

    RowLabel ws, rRev, "Revenue", True
    RowLabel ws, rGrowth, "  % growth", False
    RowLabel ws, rEbitda, "EBITDA", True
    RowLabel ws, rEbitdaM, "  % margin", False
    RowLabel ws, rDA, "Depreciation & amortization", False
    RowLabel ws, rDAPct, "  % of revenue", False
    RowLabel ws, rEbit, "EBIT", True
    RowLabel ws, rEbitM, "  % margin", False
    RowLabel ws, rTaxRate, "Tax rate", False
    RowLabel ws, rTaxes, "Less: taxes on EBIT", False
    RowLabel ws, rNopat, "NOPAT", True
    RowLabel ws, rDA2, "Plus: depreciation & amortization", False
    RowLabel ws, rCapex, "Less: capital expenditure", False
    RowLabel ws, rCapexPct, "  % of revenue", False
    RowLabel ws, rNwc, "Less: increase in net working capital", False
    RowLabel ws, rNwcPct, "  % of change in revenue", False
    RowLabel ws, rFcf, "Unlevered free cash flow", True
    RowLabel ws, rT, "Year", False
    RowLabel ws, rDiscP, "Discount period", False
    RowLabel ws, rDF, "Discount factor", False
    RowLabel ws, rPV, "PV of free cash flow", True

    ' historical columns - linked to Capital IQ
    For i = 1 To mNHist
        c = cH1 + i - 1
        y = mHistYear(i)
        PutF ws, rRev, c, "=" & PeriodLink(IT_REVENUE, y), FMT_NUM, ST_LINK
        If i > 1 Then PutF ws, rGrowth, c, "=IFERROR(" & A(rRev, c) & "/" & A(rRev, c - 1) & "-1,"""")", FMT_PCT, ST_CALC
        PutF ws, rDA, c, "=ABS(" & PeriodLink(IT_DA, y) & ")", FMT_NUM, ST_LINK
        PutF ws, rDAPct, c, "=IFERROR(" & A(rDA, c) & "/" & A(rRev, c) & ","""")", FMT_PCT, ST_CALC
        PutF ws, rEbit, c, "=" & PeriodLink(IT_EBIT, y), FMT_NUM, ST_LINK
        PutF ws, rEbitM, c, "=IFERROR(" & A(rEbit, c) & "/" & A(rRev, c) & ","""")", FMT_PCT, ST_CALC
        If mItemRow(IT_EBT) > 0 And mItemRow(IT_TAX) > 0 Then
            ebt = PeriodLink(IT_EBT, y)
            tx = PeriodLink(IT_TAX, y)
            PutF ws, rTaxRate, c, "=IFERROR(IF(" & ebt & ">0,ABS(" & tx & ")/" & ebt & ",""""),"""")", FMT_PCT, ST_LINK
        End If
        PutF ws, rTaxes, c, "=-MAX(" & A(rEbit, c) & ",0)*" & AbsA(rTax, 3), FMT_NUM, ST_CALC
        PutF ws, rCapex, c, "=-ABS(" & PeriodLink(IT_CAPEX, y) & ")", FMT_NUM, ST_LINK
        PutF ws, rCapexPct, c, "=IFERROR(-" & A(rCapex, c) & "/" & A(rRev, c) & ","""")", FMT_PCT, ST_CALC
        PutF ws, rNwc, c, "=" & PeriodLink(IT_NWC, y), FMT_NUM, ST_LINK
        If i > 1 Then
            PutF ws, rNwcPct, c, "=IFERROR(-" & A(rNwc, c) & "/(" & A(rRev, c) & "-" & A(rRev, c - 1) & "),"""")", _
                 FMT_PCT, ST_CALC
        End If
    Next i

    ' projected columns
    For j = 1 To P
        c = cP1 + j - 1
        PutF ws, rGrowth, c, "=" & AbsA(rCagr, 3) & "+(" & AbsA(rG, 3) & "-" & AbsA(rCagr, 3) & ")*" & (j - 1) & "/" & (P - 1), _
             FMT_PCT, ST_DRIVER
        PutF ws, rRev, c, "=" & A(rRev, c - 1) & "*(1+" & A(rGrowth, c) & ")", FMT_NUM, ST_CALC
        If j = 1 Then
            PutF ws, rDAPct, c, "=IFERROR(AVERAGE(" & RangeA(rDAPct, cH1, cHN) & "),0.03)", FMT_PCT, ST_DRIVER
            PutF ws, rEbitM, c, "=IFERROR(AVERAGE(" & RangeA(rEbitM, cH1, cHN) & "),0.15)", FMT_PCT, ST_DRIVER
            PutF ws, rCapexPct, c, "=IFERROR(AVERAGE(" & RangeA(rCapexPct, cH1, cHN) & "),0.04)", FMT_PCT, ST_DRIVER
            PutF ws, rNwcPct, c, "=IFERROR(MEDIAN(" & RangeA(rNwcPct, cH1, cHN) & "),0.05)", FMT_PCT, ST_DRIVER
        Else
            PutF ws, rDAPct, c, "=" & A(rDAPct, c - 1), FMT_PCT, ST_DRIVER
            PutF ws, rEbitM, c, "=" & A(rEbitM, c - 1), FMT_PCT, ST_DRIVER
            PutF ws, rCapexPct, c, "=" & A(rCapexPct, c - 1), FMT_PCT, ST_DRIVER
            PutF ws, rNwcPct, c, "=" & A(rNwcPct, c - 1), FMT_PCT, ST_DRIVER
        End If
        PutF ws, rDA, c, "=" & A(rRev, c) & "*" & A(rDAPct, c), FMT_NUM, ST_CALC
        PutF ws, rEbit, c, "=" & A(rRev, c) & "*" & A(rEbitM, c), FMT_NUM, ST_CALC
        PutF ws, rTaxRate, c, "=" & AbsA(rTax, 3), FMT_PCT, ST_DRIVER
        PutF ws, rTaxes, c, "=-MAX(" & A(rEbit, c) & ",0)*" & A(rTaxRate, c), FMT_NUM, ST_CALC
        PutF ws, rCapex, c, "=-" & A(rRev, c) & "*" & A(rCapexPct, c), FMT_NUM, ST_CALC
        PutF ws, rNwc, c, "=-(" & A(rRev, c) & "-" & A(rRev, c - 1) & ")*" & A(rNwcPct, c), FMT_NUM, ST_CALC
        ws.Cells(rT, c).Value = j
        ws.Cells(rT, c).NumberFormat = "0"
        PutF ws, rDiscP, c, "=" & A(rT, c) & "-0.5*" & AbsA(rMid, 3), "0.0", ST_CALC
        PutF ws, rDF, c, "=1/(1+" & AbsA(rWacc, 3) & ")^" & A(rDiscP, c), "0.000", ST_CALC
        PutF ws, rPV, c, "=" & A(rFcf, c) & "*" & A(rDF, c), FMT_NUM, ST_CALC
    Next j

    ' rows identical for historical and projected columns
    For c = cH1 To cPN
        PutF ws, rEbitda, c, "=" & A(rEbit, c) & "+" & A(rDA, c), FMT_NUM, ST_CALC
        PutF ws, rEbitdaM, c, "=IFERROR(" & A(rEbitda, c) & "/" & A(rRev, c) & ","""")", FMT_PCT, ST_CALC
        PutF ws, rNopat, c, "=" & A(rEbit, c) & "+" & A(rTaxes, c), FMT_NUM, ST_CALC
        PutF ws, rDA2, c, "=" & A(rDA, c), FMT_NUM, ST_CALC
        PutF ws, rFcf, c, "=" & A(rNopat, c) & "+" & A(rDA2, c) & "+" & A(rCapex, c) & "+" & A(rNwc, c), FMT_NUM, ST_CALC
    Next c
    TopBorder ws.Range(ws.Cells(rNopat, 1), ws.Cells(rNopat, cPN))
    TopBorder ws.Range(ws.Cells(rFcf, 1), ws.Cells(rFcf, cPN))
    ws.Range(ws.Cells(rFcf, 1), ws.Cells(rFcf, cPN)).Font.Bold = True
    ws.Range(ws.Cells(rPV, 1), ws.Cells(rPV, cPN)).Font.Bold = True
    For Each rowNum In Array(rGrowth, rEbitdaM, rDAPct, rEbitM, rCapexPct, rNwcPct, rTaxRate)
        ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, cPN)).Font.Italic = True
    Next rowNum

    ' ---- valuation --------------------------------------------------------------------
    Dim W As String, G As String, fcfN As String, ebitdaN As String, tN As String
    Dim ev As String, debt As String, cash As String, mi As String, pref As String, sh As String

    W = AbsA(rWacc, 3)
    G = AbsA(rG, 3)
    fcfN = AbsA(rFcf, cPN)
    ebitdaN = AbsA(rEbitda, cPN)
    tN = AbsA(rT, cPN)
    ev = AbsA(rEV, 3)
    debt = AbsA(rDebtV, 3)
    cash = AbsA(rCashV, 3)
    mi = AbsA(rMiV, 3)
    pref = AbsA(rPrefV, 3)
    sh = AbsA(rSharesV, 3)

    SectionHeader ws, rSumPV - 1, "VALUATION", cLast
    PutAssumption ws, rSumPV, "Sum of PV of free cash flows", "=SUM(" & RangeA(rPV, cP1, cPN) & ")", FMT_NUM, ST_CALC
    PutAssumption ws, rTvG, "Terminal value - Gordon growth", _
                  "=IF(" & W & ">" & G & "," & fcfN & "*(1+" & G & ")/(" & W & "-" & G & "),NA())", FMT_NUM, ST_CALC, _
                  "FCF(final year) x (1 + g) / (WACC - g)"
    PutAssumption ws, rTvX, "Terminal value - exit multiple", "=" & ebitdaN & "*" & AbsA(rMult, 3), FMT_NUM, ST_CALC, _
                  "EBITDA(final year) x exit multiple"
    PutAssumption ws, rTv, "Terminal value used", "=CHOOSE(" & AbsA(rMethod, 3) & "," & AbsA(rTvG, 3) & "," & _
                  AbsA(rTvX, 3) & ",AVERAGE(" & AbsA(rTvG, 3) & "," & AbsA(rTvX, 3) & "))", FMT_NUM, ST_CALC, _
                  "Selected with 'Terminal value method'"
    PutAssumption ws, rPvTv, "PV of terminal value", "=" & AbsA(rTv, 3) & "/(1+" & W & ")^" & tN, FMT_NUM, ST_CALC, _
                  "Discounted from the end of the final projected year"
    PutAssumption ws, rEV, "Enterprise value", "=" & AbsA(rSumPV, 3) & "+" & AbsA(rPvTv, 3), FMT_NUM, ST_CALC
    PutAssumption ws, rTvPct, "  Terminal value as % of EV", "=IFERROR(" & AbsA(rPvTv, 3) & "/" & ev & ","""")", _
                  FMT_PCT, ST_CALC
    ws.Range(ws.Cells(rEV, 1), ws.Cells(rEV, 3)).Font.Bold = True
    TopBorder ws.Range(ws.Cells(rEV, 1), ws.Cells(rEV, 3))

    PutAssumption ws, rDebtV, "Less: total debt", "=" & AbsA(rDebtA, 3), FMT_NUM, ST_CALC
    PutPoint ws, rCashV, "Plus: cash & short-term investments", IT_CASH, 0, FMT_NUM, "No cash line found - enter manually"
    PutPoint ws, rMiV, "Less: minority interest", IT_MINORITY, 0, FMT_NUM, "Not found - assumed 0"
    PutPoint ws, rPrefV, "Less: preferred equity", IT_PREFERRED, 0, FMT_NUM, "Not found - assumed 0"
    PutAssumption ws, rEqV, "Equity value", "=" & ev & "-" & debt & "+" & cash & "-" & mi & "-" & pref, FMT_NUM, ST_CALC
    PutAssumption ws, rSharesV, "Diluted shares outstanding", "=" & AbsA(rShares, 3), FMT_NUM, ST_CALC
    PutAssumption ws, rPxV, "Implied value per share", "=IFERROR(" & AbsA(rEqV, 3) & "/" & sh & ",""n/a"")", _
                  FMT_PX, ST_CALC
    PutAssumption ws, rCurPx, "Current share price", "=" & AbsA(rPrice, 3), FMT_PX, ST_CALC
    PutAssumption ws, rUpside, "Upside / (downside)", _
                  "=IFERROR(" & AbsA(rPxV, 3) & "/" & AbsA(rCurPx, 3) & "-1,""n/a"")", FMT_PCT, ST_CALC
    ws.Range(ws.Cells(rEqV, 1), ws.Cells(rEqV, 3)).Font.Bold = True
    TopBorder ws.Range(ws.Cells(rEqV, 1), ws.Cells(rEqV, 3))
    With ws.Range(ws.Cells(rPxV, 1), ws.Cells(rPxV, 3))
        .Font.Bold = True
        .Interior.Color = RGB(226, 239, 218)
    End With

    PutAssumption ws, rImpEv, "Implied EV / LTM EBITDA", "=IFERROR(" & ev & "/" & AbsA(rEbitda, cHN) & ",""n/a"")", _
                  FMT_MULT, ST_CALC, "Sanity check vs. trading comparables"
    PutAssumption ws, rImpMult, "Implied exit multiple (Gordon TV)", _
                  "=IFERROR(" & AbsA(rTvG, 3) & "/" & ebitdaN & ",""n/a"")", FMT_MULT, ST_CALC
    PutAssumption ws, rImpG, "Implied perpetual growth (exit-multiple TV)", _
                  "=IFERROR((" & AbsA(rTvX, 3) & "*" & W & "-" & fcfN & ")/(" & AbsA(rTvX, 3) & "+" & fcfN & "),""n/a"")", _
                  FMT_PCT, ST_CALC

    ' ---- sensitivities -------------------------------------------------------------
    Dim pvFcf As String, bridge As String
    pvFcf = "SUMPRODUCT(" & RangeA(rFcf, cP1, cPN, True) & ",(1+{W})^(-" & RangeA(rDiscP, cP1, cPN, True) & "))"
    bridge = "-" & debt & "+" & cash & "-" & mi & "-" & pref

    SectionHeader ws, rSens1 - 1, "SENSITIVITY - IMPLIED VALUE PER SHARE", cLast
    WriteSensitivity ws, rSens1, "WACC (down) vs. terminal growth rate (across) - Gordon growth method", _
                     W, G, SENS_G_STEP, FMT_PCT, _
                     "=IFERROR((" & pvFcf & "+" & fcfN & "*(1+{X})/({W}-{X})*(1+{W})^(-" & tN & ")" & bridge & ")/" & sh & _
                     ",""n/a"")"
    WriteSensitivity ws, rSens2, "WACC (down) vs. exit EV / EBITDA multiple (across) - exit multiple method", _
                     W, AbsA(rMult, 3), SENS_MULT_STEP, FMT_MULT, _
                     "=IFERROR((" & pvFcf & "+" & ebitdaN & "*{X}*(1+{W})^(-" & tN & ")" & bridge & ")/" & sh & _
                     ",""n/a"")"

    ' ---- mapping log ------------------------------------------------------------------
    SectionHeader ws, rMap, "CAPITAL IQ LINE-ITEM MAPPING", cLast
    ws.Cells(rMap + 1, 1).Value = "Model item"
    ws.Cells(rMap + 1, 3).Value = "Capital IQ label found"
    ws.Cells(rMap + 1, 6).Value = "Source"
    ws.Range(ws.Cells(rMap + 1, 1), ws.Cells(rMap + 1, cLast)).Font.Bold = True
    For i = 1 To IT_COUNT
        r = rMap + 1 + i
        ws.Cells(r, 1).Value = ItemName(i)
        If mItemRow(i) > 0 Then
            ws.Cells(r, 3).Value = "'" & mItemLabel(i)
            If IsPeriodItem(i) Then
                ws.Cells(r, 6).Value = "Sheet " & mSrc(mItemSheet(i)).Name & ", row " & mItemRow(i)
            Else
                ws.Cells(r, 6).Value = "Sheet " & mSrc(mItemSheet(i)).Name & ", cell " & ColL(mItemCol(i)) & mItemRow(i)
            End If
        Else
            ws.Cells(r, 3).Value = "NOT FOUND - default / 0 used"
            ws.Cells(r, 3).Font.Color = RGB(192, 0, 0)
        End If
    Next i

    ' ---- layout ---------------------------------------------------------------------
    ws.Columns(1).ColumnWidth = 42
    ws.Columns(2).ColumnWidth = 2
    ws.Range(ws.Cells(1, 3), ws.Cells(1, cLast)).EntireColumn.ColumnWidth = 12
    ws.Activate
    On Error Resume Next            ' cosmetic only
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 90
    ws.Cells(1, 1).Select
    On Error GoTo 0
End Sub

Private Sub WriteSensitivity(ByVal ws As Worksheet, ByVal top As Long, ByVal title As String, _
                             ByVal wBase As String, ByVal xBase As String, ByVal xStep As Double, _
                             ByVal xFmt As String, ByVal template As String)
    Dim i As Long, k As Long, hdr As Long, rr As Long, cc As Long, f As String

    ws.Cells(top, 1).Value = title
    ws.Cells(top, 1).Font.Bold = True
    hdr = top + 1
    ws.Cells(hdr, 3).Value = "WACC \ x"
    ws.Cells(hdr, 3).Font.Italic = True
    For k = -2 To 2
        cc = 6 + k                                  ' columns D..H
        PutF ws, hdr, cc, "=" & xBase & "+(" & k & ")*" & NumF(xStep), xFmt, ST_CALC
        ws.Cells(hdr, cc).Font.Bold = True
    Next k
    For i = -2 To 2
        rr = hdr + 3 + i
        PutF ws, rr, 3, "=" & wBase & "+(" & i & ")*" & NumF(SENS_WACC_STEP), FMT_PCT, ST_CALC
        ws.Cells(rr, 3).Font.Bold = True
        For k = -2 To 2
            cc = 6 + k
            f = Replace(template, "{W}", "$C" & rr)
            f = Replace(f, "{X}", ColL(cc) & "$" & hdr)
            PutF ws, rr, cc, f, FMT_PX, ST_CALC
        Next k
    Next i
    ws.Range(ws.Cells(hdr, 3), ws.Cells(hdr, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range(ws.Cells(hdr + 1, 3), ws.Cells(hdr + 5, 3)).Borders(xlEdgeRight).LineStyle = xlContinuous
    ws.Cells(hdr + 3, 6).Interior.Color = RGB(226, 239, 218)
End Sub

'==============================================================================
' Cell helpers
'==============================================================================

' Link expression for a period item in a given fiscal year, e.g. N('CapIQ Data'!E12).
Private Function PeriodLink(ByVal item As Long, ByVal y As Long) As String
    Dim s As Long, c As Long
    PeriodLink = "0"
    If mItemRow(item) = 0 Then Exit Function
    s = mItemSheet(item)
    c = YearCol(mData(s), s, y)
    If c > 0 Then PeriodLink = "N(" & SrcRef(s, mItemRow(item), c) & ")"
End Function

' Writes a point-in-time value (cash, debt, shares, beta, price) as a link, or a
' default input when the item was not found.
Private Sub PutPoint(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, ByVal item As Long, _
                     ByVal defaultValue As Double, ByVal fmt As String, ByVal missingNote As String)
    If mItemRow(item) > 0 And mItemCol(item) > 0 Then
        PutAssumption ws, r, label, "=N(" & SrcRef(mItemSheet(item), mItemRow(item), mItemCol(item)) & ")", fmt, ST_LINK, _
                      "Capital IQ: " & mItemLabel(item)
    Else
        PutAssumption ws, r, label, defaultValue, fmt, ST_INPUT, missingNote
        ws.Cells(r, 4).Font.Color = RGB(192, 0, 0)
    End If
End Sub

Private Sub PutAssumption(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, ByVal content As Variant, _
                          ByVal fmt As String, ByVal style As Long, Optional ByVal note As String = "")
    ws.Cells(r, 1).Value = label
    If VarType(content) = vbString Then
        PutF ws, r, 3, CStr(content), fmt, style
    Else
        ws.Cells(r, 3).Value = content
        ws.Cells(r, 3).NumberFormat = fmt
        StyleCell ws.Cells(r, 3), style
    End If
    If Len(note) > 0 Then
        ws.Cells(r, 4).Value = note
        ws.Cells(r, 4).Font.Italic = True
        ws.Cells(r, 4).Font.Color = RGB(89, 89, 89)
    End If
End Sub

Private Sub PutF(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal formula As String, _
                 ByVal fmt As String, ByVal style As Long)
    ws.Cells(r, c).Formula = formula
    ws.Cells(r, c).NumberFormat = fmt
    StyleCell ws.Cells(r, c), style
End Sub

Private Sub StyleCell(ByVal cell As Range, ByVal style As Long)
    Select Case style
        Case ST_INPUT
            cell.Font.Color = RGB(0, 0, 255)
            cell.Interior.Color = RGB(255, 242, 204)
        Case ST_DRIVER
            cell.Font.Color = RGB(0, 0, 0)
            cell.Interior.Color = RGB(255, 242, 204)
        Case ST_LINK
            cell.Font.Color = RGB(0, 128, 0)
        Case Else
            cell.Font.Color = RGB(0, 0, 0)
    End Select
End Sub

Private Sub RowLabel(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, ByVal bold As Boolean)
    ws.Cells(r, 1).Value = label
    ws.Cells(r, 1).Font.Bold = bold
End Sub

Private Sub SectionHeader(ByVal ws As Worksheet, ByVal r As Long, ByVal title As String, ByVal lastCol As Long)
    ws.Cells(r, 1).Value = title
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
End Sub

Private Sub TopBorder(ByVal rng As Range)
    rng.Borders(xlEdgeTop).LineStyle = xlContinuous
    rng.Borders(xlEdgeTop).Weight = xlThin
End Sub

' Relative (A1) and absolute ($A$1) addresses on the DCF sheet.
Private Function A(ByVal r As Long, ByVal c As Long) As String
    A = ColL(c) & r
End Function

Private Function AbsA(ByVal r As Long, ByVal c As Long) As String
    AbsA = "$" & ColL(c) & "$" & r
End Function

Private Function RangeA(ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long, Optional ByVal absolute As Boolean = False) As String
    If absolute Then
        RangeA = AbsA(r, c1) & ":" & AbsA(r, c2)
    Else
        RangeA = A(r, c1) & ":" & A(r, c2)
    End If
End Function

Private Function ColL(ByVal c As Long) As String
    Dim s As String, n As Long
    n = c
    Do While n > 0
        s = Chr$(65 + ((n - 1) Mod 26)) & s
        n = (n - 1) \ 26
    Loop
    ColL = s
End Function

Private Function SrcRef(ByVal s As Long, ByVal r As Long, ByVal c As Long) As String
    SrcRef = "'" & Replace(mSrc(s).Name, "'", "''") & "'!" & ColL(c) & r
End Function

' Locale-independent number literal for formulas (always "." as decimal separator).
Private Function NumF(ByVal x As Double) As String
    NumF = Trim$(Str$(x))
    If Left$(NumF, 1) = "." Then NumF = "0" & NumF
    If Left$(NumF, 2) = "-." Then NumF = "-0" & Mid$(NumF, 2)
End Function

'==============================================================================
' Parsing helpers
'==============================================================================

' Fiscal year in a header cell: real dates, numbers like 2023, or text such as
' "FY2023", "Dec-31-2023", "12 months" & vbLf & "Dec-31-2023", "2023A".
Private Function HeaderYear(ByVal v As Variant) As Long
    Dim s As String, i As Long, prevCh As String, nextCh As String, y As Long

    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then HeaderYear = Year(v): Exit Function
    If IsNum(v) Then
        If v >= 1980 And v <= 2100 And v = Int(v) Then HeaderYear = CLng(v)
        Exit Function
    End If
    If VarType(v) <> vbString Then Exit Function

    s = CStr(v)
    For i = 1 To Len(s) - 3
        If Mid$(s, i, 4) Like "[12][09]##" Then
            If i > 1 Then prevCh = Mid$(s, i - 1, 1) Else prevCh = ""
            nextCh = Mid$(s, i + 4, 1)
            If Not (prevCh Like "#") And Not (nextCh Like "#") Then
                y = CLng(Mid$(s, i, 4))
                If y >= 1980 And y <= 2100 Then HeaderYear = y
            End If
        End If
    Next i
End Function

' LTM, estimates, quarters, YTD etc. are not treated as historical fiscal years.
Private Function IsEstimateHeader(ByVal v As Variant) As Boolean
    Dim u As String
    If VarType(v) <> vbString Then Exit Function
    u = UCase$(Replace(Replace(CStr(v), vbLf, " "), vbCr, " "))
    u = " " & u & " "
    If InStr(u, "LTM") > 0 Or InStr(u, "NTM") > 0 Or InStr(u, "ESTIMATE") > 0 Or u Like "*[!A-Z]EST[!A-Z]*" Or _
       InStr(u, "PROJECTED") > 0 Or InStr(u, "PROJECTION") > 0 Or InStr(u, "FORECAST") > 0 Or _
       InStr(u, "YTD") > 0 Or InStr(u, "CONSENSUS") > 0 Or InStr(u, "BUDGET") > 0 Then
        IsEstimateHeader = True
    ElseIf u Like "*[12][09]##E[!A-Z]*" Then
        IsEstimateHeader = True
    ElseIf u Like "*Q[1-4]*" Or u Like "*[1-4]Q*" Or u Like "*H[12] *" Then
        IsEstimateHeader = True
    ElseIf InStr(u, "MONTHS") > 0 And InStr(u, "12 MONTHS") = 0 Then
        IsEstimateHeader = True
    End If
End Function

' Lower case, punctuation replaced by single spaces ("&" kept).
Private Function NormLabel(ByVal s As String) As String
    Dim i As Long, ch As String, out As String, lastSpace As Boolean

    s = LCase$(s)
    lastSpace = True
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[a-z0-9&]" Then
            out = out & ch
            lastSpace = False
        ElseIf Not lastSpace Then
            out = out & " "
            lastSpace = True
        End If
    Next i
    NormLabel = Trim$(out)
End Function

' Ratio / growth lines must not be picked up by the fuzzy passes.
Private Function IsExcludedLabel(ByVal s As String) As Boolean
    Dim l As String
    l = LCase$(s)
    IsExcludedLabel = (InStr(l, "%") > 0 Or InStr(l, "/") > 0 Or InStr(l, "growth") > 0 Or _
                       InStr(l, "margin") > 0 Or InStr(l, "per share") > 0 Or InStr(l, "yoy") > 0 Or _
                       InStr(l, "ratio") > 0 Or InStr(l, "multiple") > 0)
End Function

Private Function IsNum(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            IsNum = True
    End Select
End Function

Private Function IsTextCell(ByVal v As Variant) As Boolean
    If VarType(v) = vbString Then IsTextCell = (Len(Trim$(v)) > 0 And Not IsNumeric(v))
End Function

Private Function SheetExists(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, nm, vbTextCompare) = 0 Then SheetExists = True: Exit Function
    Next ws
End Function

Private Function SourceNames() As String
    Dim s As Long
    For s = 1 To mNSrc
        If mHdr(s) > 0 Or s = mItemSheet(IT_BETA) Or s = mItemSheet(IT_PRICE) Then
            If Len(SourceNames) > 0 Then SourceNames = SourceNames & ", "
            SourceNames = SourceNames & "'" & mSrc(s).Name & "'"
        End If
    Next s
End Function

Private Function MappingSummary() As String
    Dim i As Long, found As String, missing As String

    For i = 1 To IT_COUNT
        If mItemRow(i) > 0 Then
            found = found & "  " & ItemName(i) & "  <-  " & mItemLabel(i) & vbCrLf
        Else
            missing = missing & "  " & ItemName(i) & vbCrLf
        End If
    Next i
    MappingSummary = "DCF built on sheet '" & OUTPUT_SHEET & "' using FY" & mHistYear(1) & "-FY" & _
                     mHistYear(mNHist) & " as history and " & PROJ_YEARS & " projected years." & vbCrLf & vbCrLf & _
                     "Mapped from Capital IQ:" & vbCrLf & found
    If Len(missing) > 0 Then
        MappingSummary = MappingSummary & vbCrLf & "Not found (defaults / 0 used - check the yellow cells):" & _
                         vbCrLf & missing
    End If
End Function

Private Sub RaiseErr(ByVal msg As String)
    Err.Raise vbObjectError + 513, "CapIQ_DCF", msg
End Sub
