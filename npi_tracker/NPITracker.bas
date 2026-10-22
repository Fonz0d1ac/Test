Attribute VB_Name = "NPITracker"
Option Explicit

' NPI tracker update, driven by the Steps table (tblSteps on the Steps sheet).
'
' Each process step is one row in tblSteps and one pair of columns (Due date, Status) on the NPI sheet.
' The step's code sits in row 1 above its Due date column; that is how the macro finds the columns, so
' steps can be added, moved or retired without changing this code.
'   Add a step:     add a row to tblSteps, then run AddStep (inserts and formats the two columns)
'   Retire a step:  set Active = N in tblSteps (its columns are then left alone and can be deleted)
'   Check the file: run CheckLayout, which lists anything that does not line up
'
' Sheet code names used here: Sheet2 = NPI_Components, Sheet1 = Emails (the reminder log).
' Password, SCM plan location and the tracker link live in tblSettings on the Steps sheet.

' ==================== CONSTANTS ====================
Private Const START_ROW As Long = 10            ' first project row on the NPI sheet
Private Const KEY_ROW As Long = 1               ' step codes, above each step's Due date column
Private Const SECTION_ROW As Long = 7           ' PROCESS / PACKAGING / ... titles
Private Const HEADER_ROW As Long = 8            ' step titles
Private Const SUBHEADER_ROW As Long = 9         ' "Due date" / "Status"
Private Const FORMAT_LAST_ROW As Long = 4448    ' status colours and dropdowns are applied down to this row

' Project columns left of the steps (adding steps never moves them)
Private Const COL_PROJECT_NO As Long = 1        ' A
Private Const COL_PROJECT_NAME As Long = 2      ' B
Private Const COL_PART_NO As Long = 4           ' D
Private Const COL_MODE As Long = 5              ' E
Private Const COL_PB_DATE As Long = 12          ' L
Private Const COL_PB_STATUS As Long = 13        ' M
Private Const COL_SOP_DATE As Long = 14         ' N
Private Const COL_SOP_STATUS As Long = 15       ' O

Private Const STEPS_TABLE As String = "tblSteps"
Private Const SETTINGS_TABLE As String = "tblSettings"
Private Const HOLIDAYS_HEADER As String = "Holidays"
Private Const STATUS_LIST As String = "Past Due,In Progress,Ready"

Private Const NA_TEXT As String = "N/A"
Private Const NO_DATE_TEXT As String = "No Date Found"
' A row whose cell starts with this text (any case) marks the end of the data and is never processed or overwritten
Private Const END_MARKER As String = "block"

' ==================== MODULE STATE ====================
Private Type StepDef
    Code As String
    StepName As String       ' used in emails and as the column title when AddStep inserts the step
    Section As String
    IsPB As Boolean          ' due date counts from the PB sample build date (otherwise from the SOP date)
    Offset As Long           ' working days from that date; negative = before
    AutoReady As Boolean     ' set to Ready on Packaging and Raw part rows
    Active As Boolean
    Owner As String
    Email As String
    CC As String
    DueCol As Long           ' NPI sheet column of the Due date (Status is DueCol + 1); 0 = not on the sheet
End Type

Private m_Steps() As StepDef
Private m_StepCount As Long
Private m_StepIndex As Object        ' step code -> index into m_Steps
Private m_HolidayCol As Long

Private m_Phase As String            ' current step, shown in the status bar and in error messages
Private m_WorkDayCache As Object     ' WorkDay results per (base date, offset) - most rows share a few base dates
Private m_NetDaysCache As Object     ' NetworkDays results per due date
Private m_LogKeys As Object          ' notification keys already in the log (sheet rows + rows added this run)

' ==================== MAIN ENTRY POINT ====================
Public Sub Update_Date()
    Dim prevCalc As XlCalculation
    Dim prevEvents As Boolean
    Dim settingsChanged As Boolean
    Dim errs As New Collection, warns As New Collection

    On Error GoTo ErrorHandler

    Dim pwd As String
    pwd = SettingValue("Run password")
    If Len(pwd) > 0 Then
        If InputBox("Please enter password", "Password Prompt") <> pwd Then
            MsgBox "Access denied", vbCritical, "Error"
            Exit Sub
        End If
        MsgBox "Access Granted", vbInformation, "Access"
    End If

    ' Optimize performance (remember the user's settings so they can be restored exactly)
    prevCalc = Application.Calculation
    prevEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    settingsChanged = True

    SetPhase "Checking the Steps table against the NPI sheet..."
    If Not LoadLayout(errs, warns, False) Then
        ShowProblems "The update did not run. Fix these first:", errs, Nothing
        GoTo Cleanup
    End If

    SetPhase "Removing filters..."
    If Sheet1.AutoFilterMode Then Sheet1.AutoFilterMode = False
    If Sheet2.AutoFilterMode Then Sheet2.AutoFilterMode = False

    SetPhase "Updating N/A build dates from the SCM plan..."
    UpdateDatesFromSCMPlan

    ProcessAllData

    SetPhase "Finished. Restoring calculation mode (the workbook recalculates now)..."
    MsgBox "Done!", vbInformation

Cleanup:
    If settingsChanged Then
        m_Phase = "Restoring calculation mode (the workbook recalculates now)"
        Application.Calculation = prevCalc
        Application.EnableEvents = prevEvents
        Application.ScreenUpdating = True
    End If
    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    If Err.Number = 18 Then
        MsgBox "Stopped by Ctrl+Break during step: " & m_Phase, vbExclamation
    Else
        MsgBox "Error " & Err.Number & ": " & Err.Description & vbNewLine & vbNewLine & "Step: " & m_Phase, vbCritical
    End If
    Resume Cleanup
End Sub

' ==================== CHECK / ADD STEP / FORMATTING ====================
' Lists every mismatch between tblSteps, tblSettings and the NPI sheet without changing anything
Public Sub CheckLayout()
    Dim errs As New Collection, warns As New Collection
    LoadLayout errs, warns, False
    CheckSettings warns

    If errs.Count > 0 Or warns.Count > 0 Then
        ShowProblems IIf(errs.Count > 0, "The update will not run until the problems are fixed.", "No problems that stop the update."), errs, warns
        Exit Sub
    End If

    Dim i As Long, n As Long, txt As String
    For i = 1 To m_StepCount
        If m_Steps(i).Active Then
            n = n + 1
            txt = txt & vbLf & m_Steps(i).Code & "  " & ColLetter(Sheet2, m_Steps(i).DueCol) & ":" & ColLetter(Sheet2, m_Steps(i).DueCol + 1)
        End If
    Next i
    MsgBox "Everything lines up: " & n & " active steps." & vbLf & txt, vbInformation, "CheckLayout"
End Sub

' Inserts the Due date / Status columns for a step that is in tblSteps but not yet on the NPI sheet.
' The columns go right after the nearest step above it in the table (or before the nearest one below it).
Public Sub AddStep()
    Dim errs As New Collection, warns As New Collection
    Dim ws As Worksheet
    Dim i As Long, newIdx As Long, srcIdx As Long, srcDue As Long, insertCol As Long
    Dim placeAfter As Boolean, missing As String, answer As String

    If Not LoadLayout(errs, warns, True) Then
        ShowProblems "No columns were added. Fix these first:", errs, Nothing
        Exit Sub
    End If

    For i = 1 To m_StepCount
        If m_Steps(i).Active And m_Steps(i).DueCol = 0 Then
            If newIdx = 0 Then newIdx = i
            missing = missing & IIf(Len(missing) > 0, ", ", "") & m_Steps(i).Code
        End If
    Next i

    If newIdx = 0 Then
        MsgBox "Every active step in the Steps table already has its columns." & vbLf & vbLf & _
               "To add a step, first add a row to the Steps table (Code, Step name, Section, Anchor, Offset, " & _
               "Owner, Email), then run AddStep again.", vbInformation, "Add step"
        Exit Sub
    End If

    If InStr(missing, ",") > 0 Then
        answer = Trim$(InputBox("These steps have no columns yet: " & missing & vbLf & vbLf & _
                                "Which one should be added now?", "Add step", m_Steps(newIdx).Code))
        If Len(answer) = 0 Then Exit Sub
        If Not m_StepIndex.Exists(answer) Then
            MsgBox "'" & answer & "' is not a code in the Steps table.", vbExclamation, "Add step"
            Exit Sub
        End If
        newIdx = m_StepIndex(answer)
        If m_Steps(newIdx).DueCol > 0 Or Not m_Steps(newIdx).Active Then
            MsgBox "'" & answer & "' is not one of the steps waiting for columns (" & missing & ").", vbExclamation, "Add step"
            Exit Sub
        End If
    End If

    For i = newIdx - 1 To 1 Step -1
        If m_Steps(i).DueCol > 0 Then srcIdx = i: placeAfter = True: Exit For
    Next i
    If srcIdx = 0 Then
        For i = newIdx + 1 To m_StepCount
            If m_Steps(i).DueCol > 0 Then srcIdx = i: Exit For
        Next i
    End If
    If srcIdx = 0 Then
        MsgBox "No other step is on the NPI sheet, so there is no position or layout to copy.", vbExclamation, "Add step"
        Exit Sub
    End If

    Set ws = Sheet2
    srcDue = m_Steps(srcIdx).DueCol
    insertCol = IIf(placeAfter, srcDue + 2, srcDue)

    If MsgBox("Add step " & m_Steps(newIdx).Code & " (" & m_Steps(newIdx).StepName & ") as two new columns " & _
              ColLetter(ws, insertCol) & ":" & ColLetter(ws, insertCol + 1) & ", " & _
              IIf(placeAfter, "after ", "before ") & m_Steps(srcIdx).Code & "?", _
              vbOKCancel + vbQuestion, "Add step") <> vbOK Then Exit Sub

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    ws.Activate
    ws.Columns(insertCol).Resize(, 2).Insert Shift:=xlToRight
    If Not placeAfter Then srcDue = srcDue + 2

    ' Same look as the neighbouring step: merged title, date format, widths, row colours
    ws.Range(ws.Cells(HEADER_ROW, srcDue), ws.Cells(FORMAT_LAST_ROW, srcDue + 1)).Copy
    ws.Cells(HEADER_ROW, insertCol).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    ws.Columns(insertCol).ColumnWidth = ws.Columns(srcDue).ColumnWidth
    ws.Columns(insertCol + 1).ColumnWidth = ws.Columns(srcDue + 1).ColumnWidth

    Application.DisplayAlerts = False
    ws.Cells(KEY_ROW, insertCol).Value = m_Steps(newIdx).Code
    ws.Range(ws.Cells(HEADER_ROW, insertCol), ws.Cells(HEADER_ROW, insertCol + 1)).Merge
    ws.Cells(HEADER_ROW, insertCol).Value = m_Steps(newIdx).StepName
    ws.Cells(SUBHEADER_ROW, insertCol).Value = "Due date"
    ws.Cells(SUBHEADER_ROW, insertCol + 1).Value = "Status"
    PlaceInSection ws, insertCol, m_Steps(newIdx).Section
    Application.DisplayAlerts = True

    ApplyStepFormatting True
    Application.ScreenUpdating = True
    ws.Cells(HEADER_ROW, insertCol).Select

    MsgBox "Added " & m_Steps(newIdx).Code & " at " & ColLetter(ws, insertCol) & ":" & ColLetter(ws, insertCol + 1) & "." & _
           vbLf & vbLf & "Run Update_Date to fill in its due dates.", vbInformation, "Add step"
    Exit Sub

ErrorHandler:
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & " while adding the step: " & Err.Description & vbLf & vbLf & _
           "Check the NPI sheet (undo with Ctrl+Z is not available after a macro), then run CheckLayout.", vbCritical, "Add step"
End Sub

' Puts the new columns' row-7 title under the right section: joins the section on its left or right
' when the name matches, otherwise starts a new section title over the two columns.
Private Sub PlaceInSection(ws As Worksheet, newCol As Long, sectionName As String)
    Dim nb As Range, title As Variant, lastC As Long
    Dim fillNone As Boolean, fillColor As Long, fontColor As Long, isBold As Boolean

    ' Inserted inside a section: Excel has already widened its merged title
    If ws.Cells(SECTION_ROW, newCol).MergeArea.Column < newCol Then Exit Sub
    If Len(sectionName) = 0 Then Exit Sub

    Set nb = ws.Cells(SECTION_ROW, newCol - 1).MergeArea
    If SameText(nb.Cells(1, 1).Value, sectionName) Then
        ws.Range(nb.Cells(1, 1), ws.Cells(SECTION_ROW, newCol + 1)).Merge
        Exit Sub
    End If

    Set nb = ws.Cells(SECTION_ROW, newCol + 2).MergeArea
    If SameText(nb.Cells(1, 1).Value, sectionName) Then
        title = nb.Cells(1, 1).Value
        lastC = nb.Column + nb.Columns.Count - 1
        fillNone = (nb.Cells(1, 1).Interior.ColorIndex = xlColorIndexNone)
        fillColor = nb.Cells(1, 1).Interior.Color
        fontColor = nb.Cells(1, 1).Font.Color
        isBold = nb.Cells(1, 1).Font.Bold
        nb.UnMerge
        nb.Cells(1, 1).ClearContents
        With ws.Range(ws.Cells(SECTION_ROW, newCol), ws.Cells(SECTION_ROW, lastC))
            .Merge
            .Cells(1, 1).Value = title
            If fillNone Then .Interior.ColorIndex = xlColorIndexNone Else .Interior.Color = fillColor
            .Font.Color = fontColor
            .Font.Bold = isBold
            .HorizontalAlignment = xlCenter
        End With
        Exit Sub
    End If

    ws.Cells(SECTION_ROW, newCol).Value = sectionName
    ws.Range(ws.Cells(SECTION_ROW, newCol), ws.Cells(SECTION_ROW, newCol + 1)).Merge
End Sub

' Replaces the patchwork of status colour rules with one set over every status column, puts the status
' dropdown on every status column and takes it off the due-date columns
Public Sub RebuildStepFormatting()
    If MsgBox("Rebuild the status colours and dropdowns on " & Sheet2.Name & "?", vbOKCancel + vbQuestion) <> vbOK Then Exit Sub
    ApplyStepFormatting False
End Sub

Public Sub ApplyStepFormatting(Optional ByVal quiet As Boolean = False)
    Dim errs As New Collection, warns As New Collection
    Dim ws As Worksheet, statusRng As Range, dueRng As Range, block As Range, area As Range
    Dim i As Long, lastCol As Long, removed As Long

    If Not LoadLayout(errs, warns, True) Then
        ShowProblems "Formatting was not changed. Fix these first:", errs, Nothing
        Exit Sub
    End If

    Set ws = Sheet2
    lastCol = COL_SOP_STATUS
    Set statusRng = Union(BodyColumn(ws, COL_PB_STATUS), BodyColumn(ws, COL_SOP_STATUS))
    For i = 1 To m_StepCount
        If m_Steps(i).DueCol > 0 Then
            Set statusRng = Union(statusRng, BodyColumn(ws, m_Steps(i).DueCol + 1))
            If dueRng Is Nothing Then
                Set dueRng = BodyColumn(ws, m_Steps(i).DueCol)
            Else
                Set dueRng = Union(dueRng, BodyColumn(ws, m_Steps(i).DueCol))
            End If
            If m_Steps(i).DueCol + 1 > lastCol Then lastCol = m_Steps(i).DueCol + 1
        End If
    Next i

    Set block = ws.Range(ws.Cells(START_ROW, COL_PB_STATUS), ws.Cells(FORMAT_LAST_ROW, lastCol))
    For i = ws.Cells.FormatConditions.Count To 1 Step -1
        If IsStepRule(ws.Cells.FormatConditions(i), block) Then
            ws.Cells.FormatConditions(i).Delete
            removed = removed + 1
        End If
    Next i

    AddTextRule statusRng, "Past Due", RGB(255, 0, 0)
    AddTextRule statusRng, "In Progress", RGB(255, 255, 0)
    AddTextRule statusRng, "Ready", RGB(0, 176, 80)

    For Each area In statusRng.Areas
        With area.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=STATUS_LIST
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
    Next area
    If Not dueRng Is Nothing Then
        For Each area In dueRng.Areas
            area.Validation.Delete
        Next area
    End If

    If Not quiet Then
        MsgBox "Removed " & removed & " old colour rules and added 3 that cover all " & statusRng.Areas.Count & _
               " status columns.", vbInformation
    End If
End Sub

' Old status colour rules (and the due-date rule that compared against $L$1) that touch the step area
Private Function IsStepRule(fc As Object, block As Range) As Boolean
    On Error GoTo NotStepRule
    If Application.Intersect(fc.AppliesTo, block) Is Nothing Then Exit Function
    Select Case fc.Type
        Case xlTextString
            Select Case LCase$(fc.Text)
                Case "in progress", "ready", "past due": IsStepRule = True
            End Select
        Case xlCellValue
            IsStepRule = (fc.Operator = xlLess And InStr(1, fc.Formula1, "$L$1", vbTextCompare) > 0)
    End Select
NotStepRule:
End Function

Private Sub AddTextRule(rng As Range, txt As String, fillColor As Long)
    With rng.FormatConditions.Add(Type:=xlTextString, String:=txt, TextOperator:=xlContains)
        .Interior.Color = fillColor
        .StopIfTrue = False
    End With
End Sub

Private Function BodyColumn(ws As Worksheet, col As Long) As Range
    Set BodyColumn = ws.Range(ws.Cells(START_ROW, col), ws.Cells(FORMAT_LAST_ROW, col))
End Function

' ==================== LAYOUT ====================
' Reads tblSteps and finds each step's columns on the NPI sheet. Anything that would make the update wrong
' goes to errs; things worth knowing that do not stop the update go to warns. With allowMissing, an active
' step whose code is not in row 1 yet (AddStep is about to insert it) is not an error.
Private Function LoadLayout(errs As Collection, warns As Collection, allowMissing As Boolean) As Boolean
    Dim lo As ListObject
    Dim cCode As Long, cName As Long, cSection As Long, cAnchor As Long, cOffset As Long
    Dim cAuto As Long, cActive As Long, cOwner As Long, cEmail As Long, cCC As Long

    m_StepCount = 0
    Set m_StepIndex = CreateObject("Scripting.Dictionary")
    m_StepIndex.CompareMode = vbTextCompare

    Set lo = FindTable(STEPS_TABLE)
    If lo Is Nothing Then
        errs.Add "The Steps table (" & STEPS_TABLE & ") was not found. Run SetupStepsTable once to create it."
        Exit Function
    End If

    cCode = RequireColumn(lo, "Code", errs)
    cName = RequireColumn(lo, "Step name", errs)
    cSection = RequireColumn(lo, "Section", errs)
    cAnchor = RequireColumn(lo, "Anchor", errs)
    cOffset = RequireColumn(lo, "Offset", errs)
    cAuto = RequireColumn(lo, "Auto Ready (Pkg/RP)", errs)
    cActive = RequireColumn(lo, "Active", errs)
    cOwner = RequireColumn(lo, "Owner", errs)
    cEmail = RequireColumn(lo, "Email", errs)
    cCC = RequireColumn(lo, "CC", errs)
    If errs.Count > 0 Then Exit Function

    If lo.DataBodyRange Is Nothing Then
        errs.Add "The Steps table has no rows."
        Exit Function
    End If

    ' ---- Steps table ----
    Dim data As Variant, i As Long, code As String, anchor As String
    data = lo.DataBodyRange.Value2
    ReDim m_Steps(1 To UBound(data, 1))

    For i = 1 To UBound(data, 1)
        code = SafeText(data(i, cCode))
        If Len(code) = 0 Then
            If Len(SafeText(data(i, cName))) > 0 Then errs.Add "Steps table row " & i & " (" & SafeText(data(i, cName)) & ") has no Code."
        ElseIf m_StepIndex.Exists(code) Then
            errs.Add "Code " & code & " is in the Steps table twice."
        Else
            m_StepCount = m_StepCount + 1
            With m_Steps(m_StepCount)
                .Code = code
                .StepName = SafeText(data(i, cName))
                If Len(.StepName) = 0 Then .StepName = code
                .Section = SafeText(data(i, cSection))
                anchor = UCase$(SafeText(data(i, cAnchor)))
                If anchor = "PB" Or anchor = "SOP" Then
                    .IsPB = (anchor = "PB")
                Else
                    errs.Add "Step " & code & ": Anchor must be PB or SOP (it is '" & SafeText(data(i, cAnchor)) & "')."
                End If
                If IsWholeNumber(data(i, cOffset)) Then
                    .Offset = CLng(data(i, cOffset))
                Else
                    errs.Add "Step " & code & ": Offset must be a whole number of working days (it is '" & SafeText(data(i, cOffset)) & "')."
                End If
                .AutoReady = IsYes(data(i, cAuto))
                .Active = Not IsNo(data(i, cActive))
                .Owner = SafeText(data(i, cOwner))
                .Email = SafeText(data(i, cEmail))
                .CC = SafeText(data(i, cCC))
                .DueCol = 0
            End With
            m_StepIndex.Add code, m_StepCount
        End If
    Next i

    ' ---- Step codes in row 1 of the NPI sheet ----
    Dim ws As Worksheet, lastCol As Long, c As Long, k As String, idx As Long
    Dim keys As Variant, subHeads As Variant
    Set ws = Sheet2
    lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    keys = ws.Range(ws.Cells(KEY_ROW, 1), ws.Cells(KEY_ROW, lastCol + 1)).Value2
    subHeads = ws.Range(ws.Cells(SUBHEADER_ROW, 1), ws.Cells(SUBHEADER_ROW, lastCol + 1)).Value2

    For c = COL_SOP_STATUS + 1 To lastCol
        k = SafeText(keys(1, c))
        If Len(k) > 0 Then
            If Not m_StepIndex.Exists(k) Then
                errs.Add "Cell " & ColLetter(ws, c) & KEY_ROW & " has '" & k & "', which is not a code in the Steps table. " & _
                         "Add the step to the table or clear the cell."
            Else
                idx = m_StepIndex(k)
                If m_Steps(idx).DueCol > 0 Then
                    errs.Add "Code " & k & " is in row " & KEY_ROW & " twice (" & ColLetter(ws, m_Steps(idx).DueCol) & " and " & ColLetter(ws, c) & ")."
                Else
                    m_Steps(idx).DueCol = c
                    If Not SameText(subHeads(1, c), "Due date") Or Not SameText(subHeads(1, c + 1), "Status") Then
                        errs.Add "Step " & k & ": the code in " & ColLetter(ws, c) & KEY_ROW & " must sit above the step's Due date column, " & _
                                 "with its Status column right after it (row " & SUBHEADER_ROW & " should read 'Due date', 'Status')."
                    End If
                End If
            End If
        End If
    Next c

    For i = 1 To m_StepCount
        With m_Steps(i)
            If .Active And .DueCol = 0 And Not allowMissing Then
                errs.Add "Step " & .Code & " is active but its code is not in row " & KEY_ROW & " of " & ws.Name & ". " & _
                         "Run AddStep to insert its columns, or set Active = N."
            End If
            If .Active And Len(.Email) = 0 Then warns.Add "Step " & .Code & " has no Email, so its reminders cannot be sent."
        End With
    Next i

    ' ---- Fixed columns ----
    If Not SameText(ws.Cells(HEADER_ROW, COL_PB_DATE).Value, "PI/PB Sample Build date") Then _
        errs.Add "Cell " & ColLetter(ws, COL_PB_DATE) & HEADER_ROW & " should read 'PI/PB Sample Build date'."
    If Not SameText(ws.Cells(HEADER_ROW, COL_SOP_DATE).Value, "SOP date") Then _
        errs.Add "Cell " & ColLetter(ws, COL_SOP_DATE) & HEADER_ROW & " should read 'SOP date'."

    m_HolidayCol = 0
    For c = 1 To lastCol
        If SameText(subHeads(1, c), HOLIDAYS_HEADER) Then
            If m_HolidayCol = 0 Then
                m_HolidayCol = c
            Else
                errs.Add "Row " & SUBHEADER_ROW & " has '" & HOLIDAYS_HEADER & "' twice (" & ColLetter(ws, m_HolidayCol) & " and " & ColLetter(ws, c) & ")."
            End If
        End If
    Next c
    If m_HolidayCol = 0 Then errs.Add "No '" & HOLIDAYS_HEADER & "' column was found in row " & SUBHEADER_ROW & " of " & ws.Name & "."

    LoadLayout = (errs.Count = 0)
End Function

Private Sub CheckSettings(warns As Collection)
    If FindTable(SETTINGS_TABLE) Is Nothing Then
        warns.Add "The Settings table (" & SETTINGS_TABLE & ") was not found: no password, no SCM plan update, no link in emails."
        Exit Sub
    End If
    If Len(SettingValue("SCM plan file")) = 0 Then warns.Add "Setting 'SCM plan file' is blank, so N/A build dates are not filled in from the plan."
    If Len(SettingValue("SCM plan sheet")) = 0 Then warns.Add "Setting 'SCM plan sheet' is blank, so N/A build dates are not filled in from the plan."
    If Len(SettingValue("Tracker link")) = 0 Then warns.Add "Setting 'Tracker link' is blank, so emails do not link to the tracker."
    If Len(SettingValue("Run password")) = 0 Then warns.Add "Setting 'Run password' is blank, so Update_Date runs without asking for a password."
    If IsNo(SettingValue("Send emails")) Then warns.Add "Setting 'Send emails' is N, so Update_Date sends no emails (reminders stay in the log as not sent)."
End Sub

Private Function FindTable(tableName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set FindTable = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Function RequireColumn(lo As ListObject, header As String, errs As Collection) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(Trim$(lc.Name), header, vbTextCompare) = 0 Then
            RequireColumn = lc.Index
            Exit Function
        End If
    Next lc
    errs.Add "The Steps table has no '" & header & "' column."
End Function

' Value from tblSettings (column 1 = setting, column 2 = value); "" when the table or the setting is missing
Private Function SettingValue(settingName As String) As String
    Dim lo As ListObject, data As Variant, i As Long
    Set lo = FindTable(SETTINGS_TABLE)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If lo.ListColumns.Count < 2 Then Exit Function
    data = lo.DataBodyRange.Value2
    For i = 1 To UBound(data, 1)
        If SameText(data(i, 1), settingName) Then
            SettingValue = SafeText(data(i, 2))
            Exit Function
        End If
    Next i
End Function

' Shows errors and warnings in one message box (trimmed to what a message box can hold)
Private Sub ShowProblems(ByVal title As String, errs As Collection, warns As Collection)
    Dim txt As String, item As Variant, shown As Long, total As Long

    total = errs.Count
    If Not warns Is Nothing Then total = total + warns.Count
    txt = title
    If errs.Count > 0 Then
        txt = txt & vbLf & vbLf & "Problems:"
        For Each item In errs
            If Len(txt) > 850 Then Exit For
            txt = txt & vbLf & "- " & item
            shown = shown + 1
        Next item
    End If
    If Not warns Is Nothing Then
        If warns.Count > 0 And Len(txt) <= 850 Then
            txt = txt & vbLf & vbLf & "Warnings:"
            For Each item In warns
                If Len(txt) > 850 Then Exit For
                txt = txt & vbLf & "- " & item
                shown = shown + 1
            Next item
        End If
    End If
    If shown < total Then txt = txt & vbLf & vbLf & "...and " & (total - shown) & " more. Fix these and run CheckLayout again."
    MsgBox txt, IIf(errs.Count > 0, vbExclamation, vbInformation), "NPI tracker"
End Sub

' ==================== UPDATE DATES FROM SCM PLAN ====================
Private Sub UpdateDatesFromSCMPlan()
    Dim scmWb As Workbook
    Dim scmWs As Worksheet
    Dim npiWs As Worksheet
    Dim tempPath As String
    Dim prevAlerts As Boolean
    Dim planPath As String, planSheet As String

    On Error GoTo ErrorHandler
    prevAlerts = Application.DisplayAlerts

    planPath = SettingValue("SCM plan file")
    planSheet = SettingValue("SCM plan sheet")
    If Len(planPath) = 0 Or Len(planSheet) = 0 Then Exit Sub

    Set npiWs = Sheet2

    Dim npiLastRow As Long
    npiLastRow = LastDataRow(npiWs, COL_PROJECT_NO, START_ROW)
    If npiLastRow < START_ROW Then Exit Sub

    ' Load the project columns up to the SOP date in one read
    Dim npiData As Variant
    npiData = npiWs.Range(npiWs.Cells(START_ROW, 1), npiWs.Cells(npiLastRow, COL_SOP_DATE)).Value2

    ' Skip the slow network open entirely when nothing is marked N/A
    Dim i As Long
    Dim anyPending As Boolean
    For i = 1 To UBound(npiData, 1)
        If IsNA(npiData(i, COL_PB_DATE)) Or IsNA(npiData(i, COL_SOP_DATE)) Then
            anyPending = True
            Exit For
        End If
    Next i
    If Not anyPending Then Exit Sub

    ' ---- Get the plan file ----
    ' Copy it to the local temp folder first: one sequential network read is far faster and safer than
    ' letting Excel open the workbook directly over the share. Falls back to a direct open if the copy fails
    ' (for example when the file is locked by the process that writes it).
    SetPhase "SCM plan: copying file from the network..."
    tempPath = Environ$("TEMP") & "\NPV_ScmPlan_" & Format$(Now, "yyyymmdd_hhnnss") & "." & FileExtension(planPath)
    On Error Resume Next
    FileCopy planPath, tempPath
    If Err.Number <> 0 Then
        Err.Clear
        tempPath = ""
    End If
    On Error GoTo ErrorHandler

    SetPhase "SCM plan: opening file..."
    Application.DisplayAlerts = False   ' no hidden "read-only / links / recovery" prompts while the screen is frozen
    If Len(tempPath) > 0 Then
        Set scmWb = Workbooks.Open(Filename:=tempPath, ReadOnly:=True, UpdateLinks:=0, _
                                   IgnoreReadOnlyRecommended:=True, Notify:=False)
    Else
        Set scmWb = Workbooks.Open(Filename:=planPath, ReadOnly:=True, UpdateLinks:=0, _
                                   IgnoreReadOnlyRecommended:=True, Notify:=False)
    End If
    Application.DisplayAlerts = prevAlerts
    Set scmWs = scmWb.Worksheets(planSheet)

    ' ---- Find the real last data row in column B ----
    ' End(xlUp) stops at formatting-only cells and at formulas that return "", so one stray cell far down the
    ' sheet used to make the macro read (and loop over) up to a million rows. Find("*") ignores those.
    Dim lastCell As Range
    Set lastCell = scmWs.Columns("B").Find(What:="*", LookIn:=xlValues, LookAt:=xlPart, _
                                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then GoTo CleanupSCM

    Dim scmLastRow As Long
    scmLastRow = lastCell.Row
    If scmLastRow < 2 Then GoTo CleanupSCM

    ' Load SCM data (B = part number, H = delivery date, J = week info, K = PO), skipping the header row
    SetPhase "SCM plan: reading " & Format$(scmLastRow - 1, "#,##0") & " rows..."
    Dim scmData As Variant
    scmData = scmWs.Range("B2:K" & scmLastRow).Value2

    ' ---- Index the plan by part number ----
    ' One pass over the plan instead of one pass per NPI row (which was NPI rows x plan rows comparisons).
    Dim j As Long
    Dim scmPartNum As String
    Dim scmIndex As Object
    Set scmIndex = CreateObject("Scripting.Dictionary")
    scmIndex.CompareMode = vbTextCompare

    For j = 1 To UBound(scmData, 1)
        scmPartNum = SafeText(scmData(j, 1))
        If Len(scmPartNum) > 0 Then
            If scmIndex.Exists(scmPartNum) Then
                scmIndex(scmPartNum) = scmIndex(scmPartNum) & "," & j
            Else
                scmIndex.Add scmPartNum, CStr(j)
            End If
        End If
        If j Mod 5000 = 0 Then
            SetPhase "SCM plan: indexing row " & Format$(j, "#,##0") & " of " & Format$(UBound(scmData, 1), "#,##0")
            DoEvents   ' keeps Excel responsive and lets Ctrl+Break interrupt
        End If
    Next j

    ' ---- Match NPI rows ----
    Dim today As Double
    today = CDbl(VBA.Date)

    Dim npiPartNum As String
    Dim scmWeekInfo As String, poPrefix As String
    Dim deliveryDate As Double
    Dim needSample As Boolean, needMass As Boolean, isSamplePO As Boolean
    Dim bestSample As Double, bestMass As Double
    Dim rowList As Variant, k As Long

    For i = 1 To UBound(npiData, 1)
        needSample = IsNA(npiData(i, COL_PB_DATE))
        needMass = IsNA(npiData(i, COL_SOP_DATE))

        If needSample Or needMass Then
            npiPartNum = SafeText(npiData(i, COL_PART_NO))

            If Len(npiPartNum) > 0 Then
                If scmIndex.Exists(npiPartNum) Then
                    bestSample = 0
                    bestMass = 0
                    rowList = Split(scmIndex(npiPartNum), ",")

                    ' Scan every matching SCM row and keep the EARLIEST upcoming delivery
                    For k = LBound(rowList) To UBound(rowList)
                        j = CLng(rowList(k))

                        If IsRealDate(scmData(j, 7)) Then
                            deliveryDate = CDbl(scmData(j, 7))

                            If deliveryDate >= today Then
                                scmWeekInfo = SafeText(scmData(j, 9))
                                poPrefix = UCase$(Left$(SafeText(scmData(j, 10)), 2))
                                isSamplePO = (poPrefix = "PN" Or poPrefix = "PT" Or poPrefix = "PE")

                                ' Sample: week info contains "sample" AND PO starts with PN/PT/PE
                                If needSample And isSamplePO Then
                                    If InStr(1, scmWeekInfo, "sample", vbTextCompare) > 0 Then
                                        If bestSample = 0 Or deliveryDate < bestSample Then bestSample = deliveryDate
                                    End If
                                End If

                                ' Mass: week info contains "mass" AND PO does NOT start with PN/PT/PE
                                If needMass And Not isSamplePO Then
                                    If InStr(1, scmWeekInfo, "mass", vbTextCompare) > 0 Then
                                        If bestMass = 0 Or deliveryDate < bestMass Then bestMass = deliveryDate
                                    End If
                                End If
                            End If
                        End If
                    Next k

                    If needSample And bestSample > 0 Then npiWs.Cells(START_ROW + i - 1, COL_PB_DATE).Value = CDate(bestSample)
                    If needMass And bestMass > 0 Then npiWs.Cells(START_ROW + i - 1, COL_SOP_DATE).Value = CDate(bestMass)
                End If
            End If
        End If

        If i Mod 100 = 0 Then
            SetPhase "SCM plan: matching NPI row " & i & " of " & UBound(npiData, 1)
            DoEvents
        End If
    Next i

CleanupSCM:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not scmWb Is Nothing Then scmWb.Close SaveChanges:=False
    Application.DisplayAlerts = prevAlerts
    If Len(tempPath) > 0 Then Kill tempPath
    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    If Err.Number = 18 Then
        MsgBox "SCM plan update was interrupted (Ctrl+Break). Build dates were not refreshed from the plan.", vbExclamation
    Else
        MsgBox "Error updating dates from SCM Plan: " & Err.Description & vbNewLine & vbNewLine & _
               "Build dates were not refreshed from the plan. Processing will continue with the existing dates.", vbExclamation
    End If
    Resume CleanupSCM
End Sub

Private Function FileExtension(path As String) As String
    Dim dot As Long
    dot = InStrRev(path, ".")
    If dot > 0 Then FileExtension = Mid$(path, dot + 1) Else FileExtension = "xlsx"
End Function

' ==================== MAIN PROCESSING ====================
Private Sub ProcessAllData()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long, lrow3 As Long
    Dim r As Long, i As Long, rIdx As Long, cIdx As Long
    Dim arr1 As Variant, arrhl As Variant, arr2 As Variant

    SetPhase "Loading NPI data..."
    Set ws = Sheet2

    ' Per-run caches
    Set m_WorkDayCache = CreateObject("Scripting.Dictionary")
    Set m_NetDaysCache = CreateObject("Scripting.Dictionary")
    Set m_LogKeys = CreateObject("Scripting.Dictionary")
    m_LogKeys.CompareMode = vbTextCompare

    ' Column A is trimmed of trailing rows that hold no project number (for example template rows whose
    ' formula returns ""), so those rows are neither processed nor overwritten.
    lastRow = LastDataRow(ws, COL_PROJECT_NO, START_ROW)
    If lastRow < START_ROW Then
        MsgBox "No data rows found on the NPI sheet (row " & START_ROW & " onward).", vbExclamation
        Exit Sub
    End If

    lastCol = COL_SOP_STATUS
    For i = 1 To m_StepCount
        If m_Steps(i).Active And m_Steps(i).DueCol + 1 > lastCol Then lastCol = m_Steps(i).DueCol + 1
    Next i

    ' Array columns are sheet columns (the block starts at column A)
    arr1 = ws.Range(ws.Cells(START_ROW, 1), ws.Cells(lastRow, lastCol)).Value2

    ' Holidays are optional; only real dates are passed to WorkDay/NetworkDays (text there raises #VALUE!)
    arrhl = LoadHolidays(ws, m_HolidayCol, START_ROW, LastDataRow(ws, m_HolidayCol, START_ROW))

    ' Load the reminder log transposed so that ReDim Preserve can grow the row dimension
    lrow3 = LastDataRow(Sheet1, 1, 1)
    If lrow3 < 1 Then lrow3 = 1
    Dim logData As Variant
    logData = Sheet1.Range("A1:I" & lrow3).Value
    ReDim arr2(1 To 9, 1 To lrow3)
    For rIdx = 1 To lrow3
        For cIdx = 1 To 9
            arr2(cIdx, rIdx) = logData(rIdx, cIdx)
        Next cIdx
        If Len(SafeText(arr2(4, rIdx))) > 0 Then m_LogKeys(SafeText(arr2(4, rIdx))) = True
    Next rIdx

    Dim totalRows As Long
    totalRows = UBound(arr1, 1)
    For r = 1 To totalRows
        If r = 1 Or r Mod 50 = 0 Then SetPhase "Processing NPI row " & r & " of " & totalRows & "..."
        ProcessRow arr1, arrhl, arr2, r, lrow3
    Next r

    ' "Send emails" = N (for a trial run on a copy): reminders are still logged, as not sent
    If Not IsNo(SettingValue("Send emails")) Then
        SetPhase "Sending upcoming-item reminders..."
        SendUpcomingNotifications arr2
        SetPhase "Sending past-due reminders..."
        SendPastDueNotifications arr1
    End If

    ' Write back only the columns this macro fills, so nothing else on the sheet is touched
    SetPhase "Writing results back to the sheets..."
    WriteColumn ws, arr1, COL_PB_STATUS
    WriteColumn ws, arr1, COL_SOP_STATUS
    For i = 1 To m_StepCount
        If m_Steps(i).Active Then
            WriteColumn ws, arr1, m_Steps(i).DueCol
            WriteColumn ws, arr1, m_Steps(i).DueCol + 1
        End If
    Next i

    Dim outputArr As Variant
    outputArr = TransposeArray(arr2)
    Sheet1.Range("A1:I" & UBound(outputArr, 1)).Value = outputArr

    Set m_WorkDayCache = Nothing
    Set m_NetDaysCache = Nothing
    Set m_LogKeys = Nothing
End Sub

Private Sub ProcessRow(arr1 As Variant, arrhl As Variant, arr2 As Variant, r As Long, ByRef lrow3 As Long)
    Dim i As Long, anchorCol As Long

    ' Due date = anchor date +/- working days; "No Date Found" when the anchor is N/A, blank or text
    For i = 1 To m_StepCount
        With m_Steps(i)
            If .Active Then
                anchorCol = IIf(.IsPB, COL_PB_DATE, COL_SOP_DATE)
                If IsRealDate(arr1(r, anchorCol)) Then
                    arr1(r, .DueCol) = WorkDayFrom(arr1(r, anchorCol), .Offset, arrhl)
                Else
                    arr1(r, .DueCol) = NO_DATE_TEXT
                End If
            End If
        End With
    Next i

    ' Packaging and raw parts do not need the Auto Ready steps (only once at least one build date is known)
    If IsRealDate(arr1(r, COL_PB_DATE)) Or IsRealDate(arr1(r, COL_SOP_DATE)) Then
        If IsPackagingOrRawPart(arr1(r, COL_MODE)) Then
            For i = 1 To m_StepCount
                If m_Steps(i).Active And m_Steps(i).AutoReady Then arr1(r, m_Steps(i).DueCol + 1) = "Ready"
            Next i
        End If
    End If

    For i = 1 To m_StepCount
        With m_Steps(i)
            If .Active Then
                anchorCol = IIf(.IsPB, COL_PB_DATE, COL_SOP_DATE)
                If IsRealDate(arr1(r, anchorCol)) Then
                    CheckPastDue arr1, r, .DueCol, .DueCol + 1
                    CheckAndCreateNotification arr1, arr2, arrhl, r, lrow3, i
                End If
            End If
        End With
    Next i

    ' PB status covers the PB-anchored steps; SOP status covers every active step
    UpdateOverallStatus arr1, r, True, COL_PB_DATE, COL_PB_STATUS
    UpdateOverallStatus arr1, r, False, COL_SOP_DATE, COL_SOP_STATUS
End Sub

' Rows whose Mode in NPV starts with "Packaging", or is "RP" / "Raw part" (any case)
Private Function IsPackagingOrRawPart(partType As Variant) As Boolean
    Dim partStr As String
    partStr = UCase$(SafeText(partType))
    IsPackagingOrRawPart = (Left$(partStr, 9) = "PACKAGING") Or _
                           (partStr = "RP") Or _
                           (partStr = "RAW PART")
End Function

Private Sub CheckPastDue(arr1 As Variant, r As Long, dateCol As Long, statusCol As Long)
    Dim status As String
    status = SafeText(arr1(r, statusCol))
    If status = "Ready" Then Exit Sub

    If IsRealDate(arr1(r, dateCol)) Then
        If CDbl(arr1(r, dateCol)) < CDbl(Date) Then
            arr1(r, statusCol) = "Past Due"
        ElseIf status = "Past Due" Then
            ' The due date moved into the future (e.g. build date was updated): clear the stale flag
            arr1(r, statusCol) = ""
        End If
    End If
End Sub

Private Sub CheckAndCreateNotification(arr1 As Variant, arr2 As Variant, arrhl As Variant, _
                                       r As Long, ByRef lrow3 As Long, stepIdx As Long)
    Dim dateCol As Long, statusCol As Long
    dateCol = m_Steps(stepIdx).DueCol
    statusCol = dateCol + 1

    If Not IsRealDate(arr1(r, dateCol)) Then Exit Sub
    If SafeText(arr1(r, statusCol)) = "Ready" Then Exit Sub

    ' NetworkDays counts both ends, so 2 = due on the next working day
    If NetworkDaysFrom(Date, arr1(r, dateCol), arrhl) <> 2 Then Exit Sub

    Dim key As String
    key = SafeText(arr1(r, COL_PROJECT_NO)) & m_Steps(stepIdx).Code
    If NotificationExists(key) Then Exit Sub
    m_LogKeys(key) = True

    lrow3 = lrow3 + 1
    ReDim Preserve arr2(1 To 9, 1 To lrow3)

    arr2(1, lrow3) = VBA.Date
    arr2(2, lrow3) = arr1(r, COL_PROJECT_NO)
    arr2(3, lrow3) = arr1(r, COL_PROJECT_NAME)
    arr2(4, lrow3) = key                         ' Project No + item code (unique key)
    arr2(5, lrow3) = m_Steps(stepIdx).Code
    arr2(6, lrow3) = m_Steps(stepIdx).Email
    arr2(7, lrow3) = m_Steps(stepIdx).Owner
    arr2(8, lrow3) = m_Steps(stepIdx).StepName
    arr2(9, lrow3) = ""                          ' Sent status
End Sub

' True when the log (sheet rows plus rows added in this run) already has this key
Private Function NotificationExists(key As String) As Boolean
    If m_LogKeys Is Nothing Then Exit Function
    NotificationExists = m_LogKeys.Exists(key)
End Function

Private Sub UpdateOverallStatus(arr1 As Variant, r As Long, pbOnly As Boolean, dateCol As Long, statusCol As Long)
    Dim i As Long, n As Long, s As String
    Dim allReady As Boolean, allEmpty As Boolean

    allReady = True
    allEmpty = True
    For i = 1 To m_StepCount
        If m_Steps(i).Active And (m_Steps(i).IsPB Or Not pbOnly) Then
            n = n + 1
            s = SafeText(arr1(r, m_Steps(i).DueCol + 1))
            If s <> "Ready" Then allReady = False
            If s <> "" Then allEmpty = False
        End If
    Next i
    If n = 0 Then Exit Sub

    If allReady Then
        arr1(r, statusCol) = "Ready"
    ElseIf allEmpty Then
        arr1(r, statusCol) = ""
    ElseIf IsRealDate(arr1(r, dateCol)) Then
        If CDbl(arr1(r, dateCol)) < CDbl(Date) And SafeText(arr1(r, statusCol)) <> "Ready" Then
            arr1(r, statusCol) = "Past Due"
        Else
            arr1(r, statusCol) = "In Progress"
        End If
    End If
End Sub

Private Sub WriteColumn(ws As Worksheet, arr1 As Variant, col As Long)
    Dim n As Long, i As Long
    Dim out() As Variant
    n = UBound(arr1, 1)
    ReDim out(1 To n, 1 To 1)
    For i = 1 To n
        out(i, 1) = arr1(i, col)
    Next i
    ws.Cells(START_ROW, col).Resize(n, 1).Value = out
End Sub

' ==================== EMAIL NOTIFICATIONS ====================
Private Sub SendUpcomingNotifications(arr2 As Variant)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim m As Long
    Dim email As String
    Dim entry As Object

    For m = 1 To UBound(arr2, 2)
        If IsPendingLogRow(arr2, m) Then
            email = SafeText(arr2(6, m))
            ' No address (step has no Email): leave the row pending so it is retried next run
            If Len(email) > 0 Then
                Set entry = GroupFor(dict, SafeText(arr2(7, m)), email, StepCC(SafeText(arr2(5, m))))
                entry("Items") = entry("Items") & _
                    "<tr><td>" & SafeText(arr2(2, m)) & "</td><td>" & SafeText(arr2(3, m)) & "</td><td>" & SafeText(arr2(8, m)) & "</td></tr>"
                entry("Rows") = entry("Rows") & m & ","
            End If
        End If
    Next m

    If dict.Count = 0 Then Exit Sub

    SendGroupedEmails dict, "NPI Tracker Reminder", _
                      "pending items. Please review and complete them by the end of the day tomorrow.", "Item"

    ' Only mark rows as Sent when their email actually went out
    Dim key As Variant, rowList As Variant, k As Long
    For Each key In dict.Keys
        Set entry = dict(key)
        If entry("Sent") Then
            rowList = Split(entry("Rows"), ",")
            For k = LBound(rowList) To UBound(rowList)
                If Len(rowList(k)) > 0 Then arr2(9, CLng(rowList(k))) = "Sent"
            Next k
        End If
    Next key
End Sub

' A log row that has a key, has not been sent, and is not the header row
Private Function IsPendingLogRow(arr2 As Variant, m As Long) As Boolean
    If m = 1 And Not IsRealDate(arr2(1, 1)) Then Exit Function      ' header row
    If SafeText(arr2(4, m)) = "" Then Exit Function                 ' blank row
    If SafeText(arr2(9, m)) <> "" Then Exit Function                ' already sent
    IsPendingLogRow = True
End Function

Private Sub SendPastDueNotifications(arr1 As Variant)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim r As Long, i As Long
    Dim projectNo As String, projectName As String
    Dim entry As Object

    For r = 1 To UBound(arr1, 1)
        projectNo = SafeText(arr1(r, COL_PROJECT_NO))
        projectName = SafeText(arr1(r, COL_PROJECT_NAME))

        For i = 1 To m_StepCount
            With m_Steps(i)
                If .Active And Len(.Email) > 0 Then
                    If SafeText(arr1(r, .DueCol + 1)) = "Past Due" Then
                        Set entry = GroupFor(dict, .Owner, .Email, .CC)
                        entry("Items") = entry("Items") & _
                            "<tr><td>" & projectNo & "</td><td>" & projectName & "</td><td>" & .StepName & "</td></tr>"
                    End If
                End If
            End With
        Next i
    Next r

    If dict.Count > 0 Then
        SendGroupedEmails dict, "NPI Tracker - Past Due Items", "Past Due NPI items", "Past Due Item"
    End If
End Sub

' One email per recipient (To + CC). Grouping by owner name alone sent two people who share a team name
' (e.g. "ME/PE Team") a single email addressed to whichever of them came first.
Private Function GroupFor(dict As Object, owner As String, email As String, cc As String) As Object
    Dim key As String, entry As Object
    key = email & "|" & cc
    If Not dict.Exists(key) Then
        Set entry = CreateObject("Scripting.Dictionary")
        entry("Greeting") = IIf(Len(owner) > 0, owner, email)
        entry("Email") = email
        entry("CC") = cc
        entry("Items") = ""
        entry("Rows") = ""
        dict.Add key, entry
    End If
    Set GroupFor = dict(key)
End Function

Private Function StepCC(code As String) As String
    If m_StepIndex Is Nothing Then Exit Function
    If m_StepIndex.Exists(code) Then StepCC = m_Steps(m_StepIndex(code)).CC
End Function

' Sends one email per dictionary key and records entry("Sent") = True/False for each
Private Sub SendGroupedEmails(dict As Object, subject As String, itemDescription As String, itemHeader As String)
    Dim outlookApp As Object

    On Error Resume Next
    Set outlookApp = GetObject(, "Outlook.Application")
    If outlookApp Is Nothing Then Set outlookApp = CreateObject("Outlook.Application")
    On Error GoTo 0

    Dim key As Variant
    Dim entry As Object
    Dim outlookMail As Object
    Dim html As String, link As String

    If outlookApp Is Nothing Then
        For Each key In dict.Keys
            Set entry = dict(key)
            entry("Sent") = False
        Next key
        MsgBox "Outlook could not be started. No '" & subject & "' emails were sent.", vbExclamation
        Exit Sub
    End If

    link = SettingValue("Tracker link")

    Dim idx As Long
    For Each key In dict.Keys
        idx = idx + 1
        Set entry = dict(key)
        SetPhase "Sending '" & subject & "' " & idx & " of " & dict.Count & " to " & entry("Email") & "..."
        entry("Sent") = False

        html = "<font size='3' face='Calibri' color='black'>" & _
               "Dear " & entry("Greeting") & ",<br><br>" & _
               "Below are your " & itemDescription & "<br><br>" & _
               "<table border='1' cellpadding='5' cellspacing='0'>" & _
               "<tr><th>Project #</th><th>Project Name</th><th>" & itemHeader & "</th></tr>" & _
               entry("Items") & _
               "</table><br><br>"
        If Len(link) > 0 Then html = html & "Click <a href=""" & link & """>Components NPI Tracking</a> to access the file.<br><br>"
        html = html & "Best regards,<br><br>" & _
               "NPV NPI Tracker</font>"

        On Error Resume Next
        Set outlookMail = outlookApp.CreateItem(0)
        With outlookMail
            .BodyFormat = 2
            .htmlBody = html
            .To = entry("Email")
            .CC = entry("CC")
            .subject = subject
            .Send
        End With
        If Err.Number = 0 Then
            entry("Sent") = True
        Else
            Err.Clear
        End If
        On Error GoTo 0
    Next key
End Sub

' ==================== UTILITY ====================
' Last row of the contiguous block starting at firstRow (same as Ctrl+Down), returning firstRow - 1 when the
' first cell is empty. Plain End(xlDown) jumps to row 1,048,576 when the block is empty or only one row long.
Private Function LastRowFrom(ws As Worksheet, col As Long, firstRow As Long) As Long
    If IsEmpty(ws.Cells(firstRow, col).Value) Then
        LastRowFrom = firstRow - 1
    ElseIf IsEmpty(ws.Cells(firstRow + 1, col).Value) Then
        LastRowFrom = firstRow
    Else
        LastRowFrom = ws.Cells(firstRow, col).End(xlDown).Row
    End If
End Function

' Records the current step, shows it in the status bar and yields so Excel repaints and Ctrl+Break works
Private Sub SetPhase(msg As String)
    m_Phase = msg
    Application.StatusBar = msg
    DoEvents
End Sub

' LastRowFrom, then stops before the first END_MARKER row and drops trailing rows whose cell holds no text
' (blank, or a formula returning ""). Rows beyond the result are neither processed nor written back.
Private Function LastDataRow(ws As Worksheet, col As Long, firstRow As Long) As Long
    Dim blockEnd As Long
    blockEnd = LastRowFrom(ws, col, firstRow)
    LastDataRow = firstRow - 1
    If blockEnd < firstRow Then Exit Function

    Dim vals As Variant, i As Long, limit As Long
    vals = ws.Range(ws.Cells(firstRow, col), ws.Cells(blockEnd, col)).Value2
    If Not IsArray(vals) Then
        If SafeText(vals) <> "" And Not IsEndMarker(vals) Then LastDataRow = firstRow
        Exit Function
    End If

    limit = UBound(vals, 1)
    For i = 1 To UBound(vals, 1)
        If IsEndMarker(vals(i, 1)) Then
            limit = i - 1
            Exit For
        End If
    Next i

    For i = limit To 1 Step -1
        If SafeText(vals(i, 1)) <> "" Then
            LastDataRow = firstRow + i - 1
            Exit Function
        End If
    Next i
End Function

Private Function IsEndMarker(v As Variant) As Boolean
    IsEndMarker = (LCase$(Left$(SafeText(v), Len(END_MARKER))) = END_MARKER)
End Function

' Holiday list as a (n x 1) array of real dates only; Empty when there are none
Private Function LoadHolidays(ws As Worksheet, col As Long, firstRow As Long, lastRow As Long) As Variant
    LoadHolidays = Empty
    If lastRow < firstRow Then Exit Function

    Dim vals As Variant, i As Long, n As Long
    vals = ws.Range(ws.Cells(firstRow, col), ws.Cells(lastRow, col)).Value2
    If Not IsArray(vals) Then
        If IsRealDate(vals) Then LoadHolidays = vals
        Exit Function
    End If

    For i = 1 To UBound(vals, 1)
        If IsRealDate(vals(i, 1)) Then n = n + 1
    Next i
    If n = 0 Then Exit Function

    Dim out() As Variant
    ReDim out(1 To n, 1 To 1)
    n = 0
    For i = 1 To UBound(vals, 1)
        If IsRealDate(vals(i, 1)) Then
            n = n + 1
            out(n, 1) = vals(i, 1)
        End If
    Next i
    LoadHolidays = out
End Function

' True for a real date: a Date value or a positive serial number. False for text, Empty and error values.
' (IsNumeric returns False for Date values and True for Empty, so it cannot be used for this on its own.)
Private Function IsRealDate(v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    Select Case VarType(v)
        Case vbDate
            IsRealDate = True
        Case vbDouble, vbSingle, vbInteger, vbLong, vbCurrency, vbDecimal
            IsRealDate = (v > 0)
    End Select
End Function

Private Function IsWholeNumber(v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    IsWholeNumber = (CDbl(v) = Int(CDbl(v)))
End Function

' Trimmed text of a cell value; "" for Empty, Null and error values (which would otherwise raise Type mismatch)
Private Function SafeText(v As Variant) As String
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    SafeText = Trim$(CStr(v))
End Function

Private Function SameText(v As Variant, expected As String) As Boolean
    SameText = (StrComp(SafeText(v), expected, vbTextCompare) = 0)
End Function

Private Function IsYes(v As Variant) As Boolean
    Select Case UCase$(SafeText(v))
        Case "Y", "YES", "X", "TRUE", "1": IsYes = True
    End Select
End Function

Private Function IsNo(v As Variant) As Boolean
    Select Case UCase$(SafeText(v))
        Case "N", "NO", "FALSE", "0": IsNo = True
    End Select
End Function

Private Function IsNA(v As Variant) As Boolean
    IsNA = (UCase$(SafeText(v)) = NA_TEXT)
End Function

Private Function ColLetter(ws As Worksheet, col As Long) As String
    ColLetter = Split(ws.Cells(1, col).Address(True, False), "$")(0)
End Function

' WorkDay that tolerates an empty holiday list; returns a Date so General-formatted cells display it as a date.
' Results are cached per (base date, offset): thousands of rows usually share a handful of build dates, and each
' WorksheetFunction call is comparatively expensive.
Private Function WorkDayFrom(baseDate As Variant, days As Long, arrhl As Variant) As Date
    Dim key As String
    key = CStr(CDbl(baseDate)) & "|" & days

    If Not m_WorkDayCache Is Nothing Then
        If m_WorkDayCache.Exists(key) Then
            WorkDayFrom = m_WorkDayCache(key)
            Exit Function
        End If
    End If

    If IsEmpty(arrhl) Then
        WorkDayFrom = CDate(WorksheetFunction.WorkDay(baseDate, days))
    Else
        WorkDayFrom = CDate(WorksheetFunction.WorkDay(baseDate, days, arrhl))
    End If

    If Not m_WorkDayCache Is Nothing Then m_WorkDayCache.Add key, WorkDayFrom
End Function

Private Function NetworkDaysFrom(startDate As Variant, endDate As Variant, arrhl As Variant) As Long
    Dim key As String
    key = CStr(CDbl(startDate)) & "|" & CStr(CDbl(endDate))

    If Not m_NetDaysCache Is Nothing Then
        If m_NetDaysCache.Exists(key) Then
            NetworkDaysFrom = m_NetDaysCache(key)
            Exit Function
        End If
    End If

    If IsEmpty(arrhl) Then
        NetworkDaysFrom = WorksheetFunction.NetworkDays(startDate, endDate)
    Else
        NetworkDaysFrom = WorksheetFunction.NetworkDays(startDate, endDate, arrhl)
    End If

    If Not m_NetDaysCache Is Nothing Then m_NetDaysCache.Add key, NetworkDaysFrom
End Function

Private Function TransposeArray(arr As Variant) As Variant
    Dim nRows As Long, nCols As Long
    Dim i As Long, j As Long
    Dim result() As Variant

    nRows = UBound(arr, 2)
    nCols = UBound(arr, 1)

    ReDim result(1 To nRows, 1 To nCols)

    For i = 1 To nRows
        For j = 1 To nCols
            result(i, j) = arr(j, i)
        Next j
    Next i

    TransposeArray = result
End Function
