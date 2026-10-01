Option Explicit

'=====================================================================
'  MACRO 2: CREATE PIVOT REPORTS                         by SimplyData
'
'  WHAT IT DOES
'    Builds a fresh "Pivot Reports" sheet with 3 pivot tables from
'    your Report sheet. The pivots sit side by side, so they never
'    run into each other, however many rows they have.
'
'  BEFORE YOU START
'    Have a Report sheet with headers in row 1 (BuildReport fills it).
'
'  RUN IT
'    Windows Alt+F8  |  Mac Option+F8  ->  CreatePivotReports  ->  Run
'    It deletes and rebuilds "Pivot Reports" every run, so keep your
'    own notes on a different sheet.
'
'  WHAT YOU EDIT
'    PICK YOUR FIELDS (right below), plus the optional filter lines
'    and titles inside each pivot block.
'=====================================================================

Sub CreatePivotReports()

    Dim wsReport As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim eachValue As PivotField
    Dim lastRow As Long
    Dim lastCol As Long
    Dim nextCol As Long
    Dim pivotCount As Long
    Dim sourceData As String

    '====================================================
    ' PICK YOUR FIELDS
    ' Use any header from row 1 of your Report sheet.
    ' Spelling and spaces must match exactly.
    '
    '   rowField   = what you group by (the rows)
    '   valueField = what you add up or average (the values)
    '====================================================

    '--- PIVOT 1 ---
    Dim rowField1 As String, valueField1A As String, valueField1B As String
    rowField1 = "Category"
    valueField1A = "Sales"
    valueField1B = "Profit"

    '--- PIVOT 2 ---
    Dim rowField2 As String, valueField2A As String, valueField2B As String
    rowField2 = "Region"
    valueField2A = "Sales"
    valueField2B = "Profit"

    '--- PIVOT 3 ---
    Dim rowField3 As String, valueField3A As String, valueField3B As String
    rowField3 = "Ship Mode"
    valueField3A = "Sales"
    valueField3B = "Days to Ship"

    '====================================================
    ' SHEET NAME
    ' Different sheet name? Change the word in quotes.
    '====================================================
    On Error Resume Next
    Set wsReport = ThisWorkbook.Worksheets("Report")
    On Error GoTo ErrorHandler

    If wsReport Is Nothing Then
        MsgBox "Can't find the Report sheet." & vbCrLf & _
               "Run BuildReport first, or change the sheet name in quotes.", _
               vbExclamation, "Pivot Reports"
        Exit Sub
    End If

    '====================================================
    ' CHECKS (no need to edit)
    '====================================================
    lastRow = wsReport.Cells(wsReport.Rows.Count, "A").End(xlUp).Row
    lastCol = wsReport.Cells(1, wsReport.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then
        MsgBox "There is no data in the Report sheet yet." & vbCrLf & _
               "Run BuildReport first.", vbExclamation, "Pivot Reports"
        Exit Sub
    End If

    'Every field you picked must be a header in the Report sheet
    If Not PR_FieldsExist(wsReport, Array( _
            rowField1, valueField1A, valueField1B, _
            rowField2, valueField2A, valueField2B, _
            rowField3, valueField3A, valueField3B)) Then
        Exit Sub
    End If

    Application.ScreenUpdating = False

    '====================================================
    ' START A FRESH PIVOT REPORTS SHEET (no need to edit)
    '====================================================
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("Pivot Reports").Delete
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True

    Set wsPivot = ThisWorkbook.Worksheets.Add(After:=wsReport)
    wsPivot.Name = "Pivot Reports"

    '====================================================
    ' ONE DATA SOURCE FOR ALL 3 PIVOTS (no need to edit)
    '====================================================
    sourceData = wsReport.Range( _
                    wsReport.Cells(1, 1), _
                    wsReport.Cells(lastRow, lastCol) _
                 ).Address( _
                    ReferenceStyle:=xlR1C1, _
                    External:=True)

    Set pc = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=sourceData)

    '====================================================
    ' WHERE THE PIVOTS GO
    ' Each title goes in row 1 and each pivot starts in row 4,
    ' so rows 2-3 stay free for a filter dropdown.
    ' nextCol = the next free column. It's worked out for you
    ' after every pivot, so they line up side by side.
    '====================================================
    nextCol = 1

    '----------------------------------------------------
    ' PIVOT 1 - CATEGORY PERFORMANCE
    '----------------------------------------------------
    wsPivot.Cells(1, nextCol).Value = "Category Performance"

    Set pt = pc.CreatePivotTable( _
                TableDestination:=wsPivot.Cells(4, nextCol), _
                TableName:="Pivot1")

    With pt
        'Rows: group by this field
        .PivotFields(rowField1).Orientation = xlRowField

        'Values: what to add up (xlSum, xlAverage, xlCount, xlMax, xlMin)
        .AddDataField .PivotFields(valueField1A), _
                      "Total Sales", xlSum
        .AddDataField .PivotFields(valueField1B), _
                      "Total Profit", xlSum

        'OPTIONAL FILTER: delete the ' in front of line 1 to add a
        'Region dropdown. Delete it on line 2 too to pick East.
        ' .PivotFields("Region").Orientation = xlPageField
        ' .PivotFields("Region").CurrentPage = "East"
    End With

    nextCol = PR_NextFreeColumn(pt)

    '----------------------------------------------------
    ' PIVOT 2 - REGIONAL PERFORMANCE
    '----------------------------------------------------
    wsPivot.Cells(1, nextCol).Value = "Regional Performance"

    Set pt = pc.CreatePivotTable( _
                TableDestination:=wsPivot.Cells(4, nextCol), _
                TableName:="Pivot2")

    With pt
        'Rows: group by this field
        .PivotFields(rowField2).Orientation = xlRowField

        'Values: what to add up
        .AddDataField .PivotFields(valueField2A), _
                      "Total Sales", xlSum
        .AddDataField .PivotFields(valueField2B), _
                      "Total Profit", xlSum

        'OPTIONAL FILTER: delete the ' to switch it on
        ' .PivotFields("Segment").Orientation = xlPageField
        ' .PivotFields("Segment").CurrentPage = "Consumer"
    End With

    nextCol = PR_NextFreeColumn(pt)

    '----------------------------------------------------
    ' PIVOT 3 - SHIPPING PERFORMANCE
    '----------------------------------------------------
    wsPivot.Cells(1, nextCol).Value = "Shipping Performance"

    Set pt = pc.CreatePivotTable( _
                TableDestination:=wsPivot.Cells(4, nextCol), _
                TableName:="Pivot3")

    With pt
        'Rows: group by this field
        .PivotFields(rowField3).Orientation = xlRowField

        'Values: total sales, and the AVERAGE days to ship
        .AddDataField .PivotFields(valueField3A), _
                      "Total Sales", xlSum
        .AddDataField .PivotFields(valueField3B), _
                      "Average Days to Ship", xlAverage

        'OPTIONAL FILTER: delete the ' to switch it on
        ' .PivotFields("Category").Orientation = xlPageField
        ' .PivotFields("Category").CurrentPage = "Technology"
    End With

    nextCol = PR_NextFreeColumn(pt)

    '----------------------------------------------------
    ' WANT A 4TH PIVOT?
    ' Copy one whole block above (from the title line down to
    ' its nextCol line), paste it here, give it a new title and
    ' TableName:="Pivot4", and type your fields in quotes,
    ' like .PivotFields("Segment").
    '----------------------------------------------------

    '====================================================
    ' FORMAT (no need to edit)
    ' Values copy the number format from the Report sheet,
    ' so $ stays $ and % stays %.
    '====================================================
    For Each pt In wsPivot.PivotTables
        pivotCount = pivotCount + 1
        For Each eachValue In pt.DataFields
            eachValue.NumberFormat = PR_ValueFormat(wsReport, eachValue)
        Next eachValue
    Next pt

    With wsPivot.Rows(1).Font
        .Bold = True
        .Size = 14
    End With
    wsPivot.Columns.AutoFit
    wsPivot.Activate
    wsPivot.Range("A1").Select

    '====================================================
    ' FINISHED
    '====================================================
    Application.ScreenUpdating = True
    MsgBox pivotCount & " pivot tables created!", vbInformation, "Pivot Reports"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "The macro could not finish." & vbCrLf & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & vbCrLf & _
           "If you switched on a filter, check its field name and value.", _
           vbCritical, "Pivot Reports"
End Sub


'=====================================================================
' HELPERS used by CreatePivotReports. No need to change these.
'=====================================================================

' The first free column 2 columns to the right of a pivot,
' so the next pivot sits beside it with a gap.
Private Function PR_NextFreeColumn(ByVal pt As PivotTable) As Long
    PR_NextFreeColumn = pt.TableRange2.Column + pt.TableRange2.Columns.Count + 1
End Function

' Returns the column number of a header in row 1 (0 if not found).
Private Function PR_FindColumn(ByVal ws As Worksheet, ByVal fieldName As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If StrComp(CStr(ws.Cells(1, c).Value), fieldName, vbTextCompare) = 0 Then
            PR_FindColumn = c
            Exit Function
        End If
    Next c
    PR_FindColumn = 0
End Function

' Shows a clear message if a field you picked isn't a Report header.
Private Function PR_FieldsExist(ByVal ws As Worksheet, ByVal fieldNames As Variant) As Boolean
    Dim i As Long
    For i = LBound(fieldNames) To UBound(fieldNames)
        If PR_FindColumn(ws, CStr(fieldNames(i))) = 0 Then
            MsgBox "Can't find the field """ & fieldNames(i) & """ in row 1 of the " & _
                   ws.Name & " sheet." & vbCrLf & _
                   "Check the spelling in PICK YOUR FIELDS.", vbExclamation, "Pivot Reports"
            PR_FieldsExist = False
            Exit Function
        End If
    Next i
    PR_FieldsExist = True
End Function

' Picks a number format for a pivot value, based on the Report column.
Private Function PR_ValueFormat(ByVal ws As Worksheet, ByVal dataField As PivotField) As String
    Dim col As Long, sourceFormat As String

    If dataField.Function = xlCount Then
        PR_ValueFormat = "#,##0"
        Exit Function
    End If

    col = PR_FindColumn(ws, dataField.SourceName)
    If col = 0 Then
        PR_ValueFormat = "#,##0.00"
        Exit Function
    End If

    sourceFormat = ws.Cells(2, col).NumberFormat

    If dataField.Function = xlAverage And _
       (sourceFormat = "General" Or sourceFormat = "#,##0" Or sourceFormat = "0") Then
        PR_ValueFormat = "#,##0.0"
    ElseIf sourceFormat = "General" Then
        PR_ValueFormat = "#,##0.00"
    Else
        PR_ValueFormat = sourceFormat
    End If
End Function
