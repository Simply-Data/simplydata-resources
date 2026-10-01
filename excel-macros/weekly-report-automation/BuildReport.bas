Option Explicit

'=====================================================================
'  MACRO 1: BUILD REPORT                                 by SimplyData
'
'  WHAT IT DOES
'    Copies the columns you pick from the Input sheet to the Report
'    sheet, adds calculated columns, and formats everything.
'
'  BEFORE YOU START
'    1. Make 2 sheets: Input and Report
'    2. Paste your raw export into Input (headers in row 1)
'    3. Type your report headers in row 1 of Report.
'       For the practice dataset, type these in A1 to V1:
'
'       A Order ID        B Order Date      C Ship Date
'       D Ship Mode       E Customer Name   F Segment
'       G State           H Region          I Category
'       J Sub-Category    K Product Name    L Unit Price
'       M Quantity        N Discount        O Sales
'       P Profit          Q Shipping Cost   R Days to Ship
'       S Gross Sales     T Discount Amount U Profit Margin
'       V Shipping % of Sales
'
'  RUN IT
'    Windows Alt+F8  |  Mac Option+F8  ->  BuildReport  ->  Run
'    Each run ADDS the Input rows below what's already in Report,
'    so paste a fresh batch into Input each time.
'
'  WHAT YOU EDIT
'    EDIT 1: column mapping   EDIT 2: formulas   EDIT 3: formats
'    (plus the sheet names, if yours are different)
'=====================================================================

Sub BuildReport()

    Dim wsInput As Worksheet
    Dim wsReport As Worksheet
    Dim mappings As Variant
    Dim formulas As Variant
    Dim lastInputRow As Long
    Dim rowCount As Long
    Dim firstReportRow As Long
    Dim lastReportRow As Long
    Dim i As Long
    Dim sourceCol As String
    Dim reportCol As String
    Dim formulaText As String

    '====================================================
    ' SHEET NAMES
    ' Different sheet names? Change the words in quotes.
    '====================================================
    On Error Resume Next
    Set wsInput = ThisWorkbook.Worksheets("Input")
    Set wsReport = ThisWorkbook.Worksheets("Report")
    On Error GoTo ErrorHandler

    If wsInput Is Nothing Or wsReport Is Nothing Then
        MsgBox "Can't find one of your sheets." & vbCrLf & _
               "Make sure the sheet names in quotes at the top of the code " & _
               "match your sheet tabs exactly.", vbExclamation, "Build Report"
        Exit Sub
    End If

    '====================================================
    ' EDIT 1 OF 3: COLUMN MAPPING
    '
    '   Array("Input column", "Report column")
    '   First letter = from, second letter = to.
    '
    '   Add a column:    add a line, using an empty Report letter
    '   Drop a column:   delete its line
    '   Reorder:         change the 2nd letter
    '
    '   Every line ends in , _ except the last one,
    '   which has no comma.
    '====================================================
    mappings = Array( _
        Array("B", "A"), _
        Array("C", "B"), _
        Array("D", "C"), _
        Array("F", "D"), _
        Array("H", "E"), _
        Array("J", "F"), _
        Array("Q", "G"), _
        Array("S", "H"), _
        Array("W", "I"), _
        Array("Y", "J"), _
        Array("Z", "K"), _
        Array("AA", "L"), _
        Array("AC", "M"), _
        Array("AD", "N"), _
        Array("AB", "O"), _
        Array("AE", "P"), _
        Array("AF", "Q") _
    )

    '====================================================
    ' EDIT 2 OF 3: CALCULATED COLUMNS
    '
    '   Array("Report column", "formula for one row")
    '
    '   {ROW} is the row number. The macro fills it in
    '   and copies the formula all the way down.
    '   The letters point at Report columns, not Input.
    '   Text inside a formula needs double quotes: ""Late""
    '
    '   Same comma rule: every line ends in , _ except the last.
    '====================================================
    formulas = Array( _
        Array("R", "=C{ROW}-B{ROW}"), _
        Array("S", "=L{ROW}*M{ROW}"), _
        Array("T", "=S{ROW}*N{ROW}"), _
        Array("U", "=IFERROR(P{ROW}/O{ROW},0)"), _
        Array("V", "=IFERROR(Q{ROW}/O{ROW},0)") _
    )

    '====================================================
    ' COUNT YOUR ROWS
    ' "B" is the Input column it counts.
    ' Pick a column that's never blank.
    '====================================================
    lastInputRow = wsInput.Cells(wsInput.Rows.Count, "B").End(xlUp).Row

    If lastInputRow < 2 Then
        MsgBox "The Input sheet has no data yet." & vbCrLf & _
               "Paste your export under the headers in row 1.", _
               vbExclamation, "Build Report"
        Exit Sub
    End If

    rowCount = lastInputRow - 1

    '====================================================
    ' CHECKS (no need to edit)
    '====================================================
    If Trim(CStr(wsReport.Range("A1").Value)) = "" Then
        MsgBox "Type your report headers in row 1 of the Report sheet first." & _
               vbCrLf & "The list is at the top of this code.", _
               vbExclamation, "Build Report"
        Exit Sub
    End If

    'The first empty row in Report
    firstReportRow = wsReport.Cells(wsReport.Rows.Count, "A").End(xlUp).Row + 1
    lastReportRow = firstReportRow + rowCount - 1

    'One last check, so you don't add the same batch twice
    If MsgBox("Add " & rowCount & " rows from Input to Report?", _
              vbYesNo + vbQuestion, "Build Report") = vbNo Then
        Exit Sub
    End If

    Application.ScreenUpdating = False

    '====================================================
    ' COPY THE MAPPED COLUMNS (no need to edit)
    '====================================================
    For i = LBound(mappings) To UBound(mappings)
        sourceCol = mappings(i)(0)
        reportCol = mappings(i)(1)

        wsReport.Range(reportCol & firstReportRow & ":" & reportCol & lastReportRow).Value = _
            wsInput.Range(sourceCol & "2:" & sourceCol & lastInputRow).Value
    Next i

    '====================================================
    ' ADD THE CALCULATED COLUMNS (no need to edit)
    ' Excel moves the row number down for every row.
    '====================================================
    For i = LBound(formulas) To UBound(formulas)
        reportCol = formulas(i)(0)
        formulaText = Replace(formulas(i)(1), "{ROW}", firstReportRow)

        wsReport.Range(reportCol & firstReportRow & ":" & reportCol & lastReportRow).Formula = _
            formulaText
    Next i

    '====================================================
    ' EDIT 3 OF 3: FORMATS
    '
    '   wsReport.Range("which columns").NumberFormat = "which format"
    '
    '   Date          "m/d/yyyy"
    '   Money         "$#,##0.00"
    '   Percent       "0.0%"
    '   Whole number  "#,##0"
    '
    '   Added a column W? Change A:V to A:W in the AutoFit line.
    '====================================================
    wsReport.Range("B:C").NumberFormat = "m/d/yyyy"
    wsReport.Range("L:L").NumberFormat = "$#,##0.00"
    wsReport.Range("N:N").NumberFormat = "0.0%"
    wsReport.Range("O:Q").NumberFormat = "$#,##0.00"
    wsReport.Range("R:R").NumberFormat = "#,##0"
    wsReport.Range("S:T").NumberFormat = "$#,##0.00"
    wsReport.Range("U:V").NumberFormat = "0.0%"
    wsReport.Rows(1).Font.Bold = True
    wsReport.Columns("A:V").AutoFit

    '====================================================
    ' FINISHED
    '====================================================
    Application.ScreenUpdating = True
    MsgBox rowCount & " rows added to the Report.", vbInformation, "Build Report"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "The macro could not finish." & vbCrLf & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & vbCrLf & _
           "Check the letters in EDIT 1 and EDIT 2.", vbCritical, "Build Report"
End Sub
