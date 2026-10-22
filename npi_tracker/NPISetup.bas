Attribute VB_Name = "NPISetup"
Option Explicit

' One-time setup for the table-driven tracker in module NPITracker. Run SetupStepsTable once, then delete
' this module. It:
'   - adds a Steps sheet with tblSteps (today's 19 steps, contacts copied from Items_PIC) and tblSettings
'   - replaces the column numbers in row 1 of the NPI sheet with each step's code, above its Due date column
'   - points the Update button at NPITracker.Update_Date
'   - rebuilds the status colours and dropdowns
' Sheet code names used here: Sheet2 = NPI_Components, Sheet3 = Items_PIC.

Private Const STEPS_SHEET As String = "Steps"

' Code | Step name (used in emails) | Section | Anchor | Offset in working days | Auto Ready for Packaging/Raw part |
' Due date column on the NPI sheet today | title expected in row 8 of that column | Items_PIC code to copy the contacts from
Private Function DefaultSteps() As Variant
    DefaultSteps = Array( _
        "PFD|Process Flow|PROCESS|PB|-1|Y|P|PFC|PFD", _
        "TM|Workbench, Tooling and Machine|PROCESS|PB|-1|Y|R|Workbench, Tooling and Machine|TM", _
        "PRWI|Work Instruction|PROCESS|PB|-1|N|T|Work Instruction|PRWI", _
        "TSE|Time Study Estimate & Update Database|PROCESS|PB|-1|N|V|Time Study Estimate and Update MainDatabase|TSE", _
        "MHO|Equipment Specs Handover to Maintenance|PROCESS|PB|-1|Y|X|Equi. specs Handover to Maint.|MHO", _
        "TSA|Time Study Audit|PROCESS|SOP|5|N|Z|Time Study Audit|TSA", _
        "PFMEA|PFMEA|PROCESS|SOP|-5|Y|AB|PFMEA|PFMEA", _
        "PAWI|Packaging Work Instruction|PACKAGING|PB|-5|N|AD|Work Instruction|PAWI", _
        "BOME|BOM Release - Estimate|PACKAGING|PB|-1|N|AF|BOM Release - Estimate|BOME", _
        "MOQ|Input MOQ and Order Multiple|PACKAGING|SOP|10|N|AH|Input MOQ and Order Multiple|MOQ", _
        "TR|Trial Run and Review EOQ-MOQ|PACKAGING|SOP|-5|N|AJ|Trial Run and Review EOQ-MOQ|TR", _
        "BOMF|BOM Release - Fixed|PACKAGING|SOP|-10|N|AL|BOM Release - Fixed|BOMF", _
        "PIIC|Send Packaging Standard to PII|PACKAGING|SOP|10|N|AN|Send Packaging Standard to PII for confirmation|PIIC", _
        "RPSV|Receive Packing Specs from Vendor|PACKAGING|SOP|-5|N|AP|Receive Packing Specs from Vendor|DFMEA", _
        "CAPP|Capacity Planning|Supervisor|PB|-5|N|AR|Capacity Planning|CAPP", _
        "FA|1st Safety Assessment|EHS|PB|-1|Y|AT|1st assessment|FA", _
        "SA|2nd Safety Assessment|EHS|SOP|-1|Y|AV|2nd assessment|SA", _
        "PRM|Preventative Maintenance|Maintenance|SOP|5|Y|AX|Preventative Maintenance|PRM", _
        "SP|Spare Parts|Maintenance|SOP|5|Y|AZ|Spare parts|SP")
End Function

Public Sub SetupStepsTable()
    Dim npi As Worksheet, st As Worksheet, lo As ListObject, sh As Shape
    Dim steps As Variant, f As Variant, contacts As Object, c As Variant
    Dim i As Long, r As Long, problems As String
    Dim planPath As String, link As String, pwd As String

    Set npi = Sheet2

    If TableExists("tblSteps") Then
        MsgBox "The Steps table already exists, so setup has already run. Use CheckLayout to check the file.", vbInformation
        Exit Sub
    End If
    If SheetExists(STEPS_SHEET) Then
        MsgBox "A sheet named '" & STEPS_SHEET & "' already exists. Rename or delete it, then run setup again.", vbExclamation
        Exit Sub
    End If

    ' Only go ahead when the NPI sheet still has the layout these defaults describe
    steps = DefaultSteps()
    If Not SameText(npi.Range("L8").Value, "PI/PB Sample Build date") Then problems = problems & vbLf & "L8 should read 'PI/PB Sample Build date'"
    If Not SameText(npi.Range("N8").Value, "SOP date") Then problems = problems & vbLf & "N8 should read 'SOP date'"
    For i = LBound(steps) To UBound(steps)
        f = Split(steps(i), "|")
        If Not SameText(npi.Range(f(6) & "8").Value, CStr(f(7))) Then
            problems = problems & vbLf & f(6) & "8 should read '" & f(7) & "' but reads '" & npi.Range(f(6) & "8").Text & "'"
        End If
        If Not SameText(npi.Range(f(6) & "9").Value, "Due date") Or _
           Not SameText(npi.Range(f(6) & "9").Offset(0, 1).Value, "Status") Then
            problems = problems & vbLf & f(6) & "9 should read 'Due date', with 'Status' next to it"
        End If
    Next i
    If Len(problems) > 0 Then
        MsgBox "Setup stopped: " & npi.Name & " is not laid out the way this setup expects." & vbLf & problems, vbExclamation
        Exit Sub
    End If

    If MsgBox("This will:" & vbLf & _
              "- add a '" & STEPS_SHEET & "' sheet with the 19 steps (contacts copied from " & Sheet3.Name & ")" & vbLf & _
              "- replace the column numbers in row 1 of " & npi.Name & " (L1:BA1) with the step codes" & vbLf & _
              "- point the Update button at NPITracker.Update_Date" & vbLf & _
              "- rebuild the status colours and dropdowns" & vbLf & vbLf & _
              "DFMEA becomes RPSV (Receive Packing Specs from Vendor)." & vbLf & vbLf & _
              "Save a copy of the workbook first. Continue?", vbOKCancel + vbQuestion, "NPI tracker setup") <> vbOK Then Exit Sub

    ' Settings that used to be written in the old macro (module UpdateDateV2)
    planPath = InputBox("SCM plan file: paste the value of SCM_PLAN_PATH from the old module UpdateDateV2." & vbLf & vbLf & _
                        "Leave blank to skip filling N/A build dates from the plan.", "Setting 1 of 3")
    link = InputBox("Tracker link for the emails: paste the https:// address after 'Click <a href=' in the old " & _
                    "module UpdateDateV2 (without the quotes)." & vbLf & vbLf & "Leave blank to leave the link out.", "Setting 2 of 3")
    pwd = InputBox("Password for Update_Date: the value of PASSWORD in the old module UpdateDateV2." & vbLf & vbLf & _
                   "Leave blank to run without a password.", "Setting 3 of 3")

    Application.ScreenUpdating = False

    ' ---- Contacts from Items_PIC (B = code, E = name, F = email, G = CC) ----
    Set contacts = CreateObject("Scripting.Dictionary")
    contacts.CompareMode = vbTextCompare
    With Sheet3
        For r = 2 To .Cells(.Rows.Count, "B").End(xlUp).Row
            If Len(Trim$(CStr(.Cells(r, "B").Value))) > 0 And Not contacts.Exists(Trim$(CStr(.Cells(r, "B").Value))) Then
                contacts.Add Trim$(CStr(.Cells(r, "B").Value)), _
                    Array(Trim$(CStr(.Cells(r, "E").Value)), Trim$(CStr(.Cells(r, "F").Value)), Trim$(CStr(.Cells(r, "G").Value)))
            End If
        Next r
    End With

    ' ---- Steps sheet and tblSteps ----
    Set st = ThisWorkbook.Worksheets.Add(After:=Sheet3)
    st.Name = STEPS_SHEET
    st.Range("A1:L1").Value = Array("Code", "Step name", "Section", "Anchor", "Offset", "Auto Ready (Pkg/RP)", _
                                    "Active", "Owner", "Email", "CC", "Due rule", "Past due")
    For i = LBound(steps) To UBound(steps)
        f = Split(steps(i), "|")
        r = i - LBound(steps) + 2
        st.Cells(r, 1).Value = f(0)
        st.Cells(r, 2).Value = f(1)
        st.Cells(r, 3).Value = f(2)
        st.Cells(r, 4).Value = f(3)
        st.Cells(r, 5).Value = CLng(f(4))
        st.Cells(r, 6).Value = f(5)
        st.Cells(r, 7).Value = "Y"
        If contacts.Exists(CStr(f(8))) Then
            c = contacts(CStr(f(8)))
            st.Cells(r, 8).Value = c(0)
            st.Cells(r, 9).Value = c(1)
            st.Cells(r, 10).Value = c(2)
        End If
    Next i

    Set lo = st.ListObjects.Add(xlSrcRange, st.Range("A1").Resize(UBound(steps) - LBound(steps) + 2, 12), , xlYes)
    lo.Name = "tblSteps"
    lo.TableStyle = "TableStyleMedium2"
    lo.ListColumns("Offset").DataBodyRange.NumberFormat = "+0;-0;0"
    lo.ListColumns("Due rule").DataBodyRange.Formula = _
        "=IF([@Offset]=0,""On the "",ABS([@Offset])&"" working day""&IF(ABS([@Offset])=1,"""",""s"")&IF([@Offset]<0,"" before "","" after ""))&[@Anchor]&"" date"""
    lo.ListColumns("Past due").DataBodyRange.Formula = _
        "=IFERROR(COUNTIF(INDEX('" & npi.Name & "'!$A:$XFD,0,MATCH([@Code],'" & npi.Name & "'!$1:$1,0)+1),""Past Due""),0)"
    AddListValidation lo.ListColumns("Anchor").DataBodyRange, "PB,SOP"
    AddListValidation lo.ListColumns("Auto Ready (Pkg/RP)").DataBodyRange, "Y,N"
    AddListValidation lo.ListColumns("Active").DataBodyRange, "Y,N"

    ' ---- tblSettings ----
    st.Range("N1:O1").Value = Array("Setting", "Value")
    st.Range("N2").Value = "Run password"
    st.Range("O2").NumberFormat = "@"
    st.Range("O2").Value = pwd
    st.Range("N3").Value = "SCM plan file"
    st.Range("O3").Value = planPath
    st.Range("N4").Value = "SCM plan sheet"
    st.Range("O4").Value = "ScmPlan"
    st.Range("N5").Value = "Tracker link"
    st.Range("O5").Value = link
    st.Range("N6").Value = "Send emails"
    st.Range("O6").Value = "Y"
    Set lo = st.ListObjects.Add(xlSrcRange, st.Range("N1:O6"), , xlYes)
    lo.Name = "tblSettings"
    lo.TableStyle = "TableStyleLight9"
    AddListValidation st.Range("O6"), "Y,N"

    st.Range("N8").Value = "How to change the steps"
    st.Range("N8").Font.Bold = True
    st.Range("N9").Value = "Add a step: add a row to the table on the left, then run the macro AddStep."
    st.Range("N10").Value = "Retire a step: set Active = N (its columns are then ignored and can be deleted)."
    st.Range("N11").Value = "Change a due date rule, owner or email: edit the row; the next Update_Date uses it."
    st.Range("N12").Value = "Row 1 of " & npi.Name & " holds each step's code above its Due date column. Do not move or retype these codes."
    st.Range("N13").Value = "Run the macro CheckLayout any time to check that the table and the sheet still line up."

    st.Columns("A:L").AutoFit
    st.Columns("B").ColumnWidth = 40
    st.Columns("N").ColumnWidth = 18
    st.Columns("O").ColumnWidth = 60

    ' ---- Step codes in row 1 of the NPI sheet ----
    npi.Range("L1:BA1").ClearContents
    For i = LBound(steps) To UBound(steps)
        f = Split(steps(i), "|")
        npi.Range(f(6) & "1").Value = f(0)
    Next i

    ' ---- Button ----
    On Error Resume Next
    For Each sh In npi.Shapes
        If InStr(1, sh.OnAction, "Update_Date", vbTextCompare) > 0 Then sh.OnAction = "NPITracker.Update_Date"
    Next sh
    On Error GoTo 0

    NPITracker.ApplyStepFormatting True

    Application.ScreenUpdating = True
    st.Activate
    st.Range("A1").Select

    MsgBox "Setup done." & vbLf & vbLf & _
           "Next:" & vbLf & _
           "1. Check the Owner / Email / CC columns on the " & STEPS_SHEET & " sheet (copied from " & Sheet3.Name & ")." & vbLf & _
           "2. Run CheckLayout." & vbLf & _
           "3. Delete the old modules Old, OldV2, Module1, UpdateDateV2 and this module, NPISetup." & vbLf & _
           "4. Run Update_Date as usual.", vbInformation, "NPI tracker setup"
End Sub

Private Sub AddListValidation(rng As Range, items As String)
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=items
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

Private Function TableExists(tableName As String) As Boolean
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then TableExists = True
        Next lo
    Next ws
End Function

Private Function SheetExists(sheetName As String) As Boolean
    Dim ws As Object
    For Each ws In ThisWorkbook.Sheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then SheetExists = True
    Next ws
End Function

Private Function SameText(v As Variant, expected As String) As Boolean
    If IsError(v) Then Exit Function
    SameText = (StrComp(Trim$(CStr(v)), expected, vbTextCompare) = 0)
End Function
