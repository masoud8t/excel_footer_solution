' =============================================================================
' KISS Print - Excel Bottom-of-Page Repeater
'
' Select the rows that should repeat at the bottom of each page,
' then run KISS_Print_Form.
'
' The generated print sheet is intentionally KEPT for manual adjustment.
' =============================================================================

Option Explicit

Public Sub KISS_Print_Form()

    Const DATA_FIRST_ROW As Long = 8
    Const FOOTER_GAP_ROWS As Long = 1

    Dim src As Worksheet
    Dim tmp As Worksheet

    Dim footerFirst As Long
    Dim footerLast As Long
    Dim footerCount As Long

    Dim rowsPerPage As Long
    Dim userInput As Variant

    Dim printFirstCol As Long
    Dim printLastCol As Long
    Dim printFirstRow As Long

    Dim dataLast As Long
    Dim pageEnd As Long

    Dim insertAt As Long
    Dim footerDest As Long
    Dim footerEnd As Long

    Dim i As Long
    Dim sheetName As String

    Dim oldAlerts As Boolean
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean

    On Error GoTo Fail

    Set src = ActiveSheet

    ' -------------------------------------------------------------------------
    ' Footer = currently selected rows
    ' -------------------------------------------------------------------------
    If Selection Is Nothing Then
        MsgBox "Select the footer rows first.", _
               vbExclamation, "KISS Print"
        Exit Sub
    End If

    footerFirst = Selection.Row
    footerLast = Selection.Row + Selection.Rows.Count - 1
    footerCount = footerLast - footerFirst + 1

    ' -------------------------------------------------------------------------
    ' Ask how many data rows should appear before the footer.
    ' -------------------------------------------------------------------------
    userInput = Application.InputBox( _
        Prompt:="After how many data rows should the selected footer be pasted?", _
        Title:="KISS Print - Rows Per Page", _
        Default:=24, _
        Type:=1)

    If VarType(userInput) = vbBoolean Then
        If userInput = False Then Exit Sub
    End If

    If userInput < 1 Or userInput <> Int(userInput) Then
        MsgBox "Please enter a positive whole number.", _
               vbExclamation, "KISS Print"
        Exit Sub
    End If

    rowsPerPage = CLng(userInput)

    ' -------------------------------------------------------------------------
    ' Determine the source print area.
    ' -------------------------------------------------------------------------
    Dim printRng As Range

    If Len(src.PageSetup.PrintArea) > 0 Then
        Set printRng = src.Range(src.PageSetup.PrintArea)
    Else
        Set printRng = src.UsedRange
    End If

    printFirstRow = printRng.Row
    printFirstCol = printRng.Column
    printLastCol = printRng.Column + printRng.Columns.Count - 1

    dataLast = footerFirst - 1

    If DATA_FIRST_ROW > dataLast Then
        MsgBox "DATA_FIRST_ROW is below the footer.", _
               vbCritical, "KISS Print"
        Exit Sub
    End If

    ' -------------------------------------------------------------------------
    ' Save Excel state.
    ' -------------------------------------------------------------------------
    oldAlerts = Application.DisplayAlerts
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' -------------------------------------------------------------------------
    ' Create a new copy. Existing KISS print sheets are preserved.
    ' -------------------------------------------------------------------------
    src.Copy After:=src.Parent.Worksheets(src.Parent.Worksheets.Count)

    Set tmp = src.Parent.Worksheets(src.Parent.Worksheets.Count)

    sheetName = GetUniqueSheetName(src.Parent, "__KISS_PRINT")
    tmp.Name = sheetName

    ' -------------------------------------------------------------------------
    ' Remove the original footer from the working copy.
    ' -------------------------------------------------------------------------
    tmp.Rows(footerFirst & ":" & footerLast).Delete

    tmp.ResetAllPageBreaks

    ' -------------------------------------------------------------------------
    ' Build pages.
    '
    ' This is intentionally simple:
    '   N data rows
    '   repeated footer
    '   page break
    '
    ' The user controls N.
    ' -------------------------------------------------------------------------
    pageEnd = dataLast

    Do While pageEnd >= DATA_FIRST_ROW

        insertAt = pageEnd + 1

        ' Reserve:
        '   spacer row(s)
        '   footer rows
        tmp.Rows( _
            insertAt & ":" & _
            insertAt + FOOTER_GAP_ROWS + footerCount - 1 _
        ).Insert Shift:=xlDown

        footerDest = insertAt + FOOTER_GAP_ROWS

        ' Copy the selected footer from the original worksheet.
        src.Range( _
            src.Cells(footerFirst, printFirstCol), _
            src.Cells(footerLast, printLastCol) _
        ).Copy Destination:= _
            tmp.Cells(footerDest, printFirstCol)

        ' Preserve row heights.
        For i = 0 To footerCount - 1
            tmp.Rows(footerDest + i).RowHeight = _
                src.Rows(footerFirst + i).RowHeight
        Next i

        footerEnd = footerDest + footerCount - 1

        ' Page break after footer.
        If pageEnd - rowsPerPage >= DATA_FIRST_ROW Then
            tmp.HPageBreaks.Add _
                Before:=tmp.Rows(footerEnd + 1)
        End If

        pageEnd = pageEnd - rowsPerPage

    Loop

    ' -------------------------------------------------------------------------
    ' Set the print area.
    ' -------------------------------------------------------------------------
    Dim finalLastRow As Long

    finalLastRow = tmp.UsedRange.Row + _
                   tmp.UsedRange.Rows.Count - 1

    tmp.PageSetup.PrintArea = _
        tmp.Range( _
            tmp.Cells(printFirstRow, printFirstCol), _
            tmp.Cells(finalLastRow, printLastCol) _
        ).Address

    ' -------------------------------------------------------------------------
    ' Fit horizontally to one page; leave vertical pagination alone.
    ' -------------------------------------------------------------------------
    On Error Resume Next

    Application.PrintCommunication = False

    With tmp.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With

    Application.PrintCommunication = True

    On Error GoTo Fail

    ' -------------------------------------------------------------------------
    ' Preview and KEEP the generated worksheet.
    ' -------------------------------------------------------------------------
    Application.ScreenUpdating = True

    tmp.Activate
    tmp.PrintPreview

    Application.DisplayAlerts = oldAlerts
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents

    MsgBox _
        "Print sheet created and kept:" & vbCrLf & vbCrLf & _
        tmp.Name & vbCrLf & vbCrLf & _
        "Data rows per page: " & rowsPerPage & vbCrLf & _
        "Footer rows: " & footerCount & vbCrLf & vbCrLf & _
        "You can now manually adjust the print sheet.", _
        vbInformation, _
        "KISS Print"

    Exit Sub

Fail:

    On Error Resume Next

    Application.PrintCommunication = True
    Application.DisplayAlerts = oldAlerts
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents

    On Error GoTo 0

    MsgBox _
        "KISS Print failed:" & vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "KISS Print"

End Sub

Private Function GetUniqueSheetName( _
    ByVal wb As Workbook, _
    ByVal baseName As String) As String

    Dim i As Long
    Dim testName As String

    testName = baseName

    If Not SheetExists(wb, testName) Then
        GetUniqueSheetName = testName
        Exit Function
    End If

    i = 2

    Do
        testName = baseName & "_" & i

        If Not SheetExists(wb, testName) Then
            GetUniqueSheetName = testName
            Exit Function
        End If

        i = i + 1
    Loop

End Function

Private Function SheetExists( _
    ByVal wb As Workbook, _
    ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = wb.Worksheets(sheetName)

    SheetExists = Not ws Is Nothing

    On Error GoTo 0

End Function
