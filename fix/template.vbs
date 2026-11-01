' excel_qa_fix.vbs - apply approved fixes with REAL Excel via late binding (ASCII only)
' Generated per job by Claude from the excel-qa skill template. Double-clicked through RUN_FIX_<base>.bat.
' Source is opened READ-ONLY (macros off, links not updated) and saved as a NEW file. Self-deletes at the end.
' Non-ASCII text (Thai) is written as \uXXXX and decoded with U().
Option Explicit
Dim fso, here, src, dst, errFile, logTxt, xl, wb, orig, pwRequired, gCount
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
'@@CONFIG@@
errFile = here & "\excel_qa_ERROR.txt"
logTxt = ""

Function U(t)
  Dim k, r
  r = "" : k = 1
  Do While k <= Len(t)
    If Mid(t, k, 2) = "\u" Then
      r = r & ChrW(CLng("&H" & Mid(t, k + 2, 4))) : k = k + 6
    Else
      r = r & Mid(t, k, 1) : k = k + 1
    End If
  Loop
  U = r
End Function

Sub Note(m)
  logTxt = logTxt & m & vbCrLf
  WScript.Echo m
End Sub

Sub Finish(code)
  On Error Resume Next
  If IsObject(wb) Then wb.Close False
  If IsObject(xl) Then xl.Quit
  Set wb = Nothing : Set xl = Nothing
  If code <> 0 Then
    Dim tf
    Set tf = fso.CreateTextFile(errFile, True, True)
    tf.Write logTxt
    tf.Close
  End If
  fso.DeleteFile WScript.ScriptFullName, True
  WScript.Quit code
End Sub

Sub Fail(m)
  Note "ERROR: " & m
  Finish 1
End Sub

Sub Chk(stepName)
  If Err.Number <> 0 Then Fail stepName & ": " & Err.Description & " (0x" & Hex(Err.Number) & ")"
End Sub

Function Sh(name)
  On Error Resume Next
  Set Sh = wb.Worksheets(U(name))
  Chk "sheet not found: " & name
End Function

' ---------- operations (call only the ones the user approved) ----------
Sub SetF(sheetName, addr, txt)          ' formula or constant text/number
  On Error Resume Next
  Sh(sheetName).Range(addr).Formula = U(txt)
  Chk "set " & sheetName & "!" & addr
  gCount = gCount + 1
End Sub

Sub SetFormat(sheetName, addr, fmt)
  On Error Resume Next
  Sh(sheetName).Range(addr).NumberFormat = U(fmt)
  Chk "format " & sheetName & "!" & addr
End Sub

Sub SetDate(sheetName, addr, y, m, d, fmt) ' fmt "" = keep current format
  On Error Resume Next
  Dim r
  Set r = Sh(sheetName).Range(addr)
  If fmt <> "" Then r.NumberFormat = U(fmt)
  r.Value2 = CDbl(DateSerial(y, m, d))
  Chk "date " & sheetName & "!" & addr
End Sub

Sub SheetUnhide(sheetName)
  On Error Resume Next
  Sh(sheetName).Visible = -1
  Chk "unhide " & sheetName
End Sub

Sub SheetHide(sheetName, very)          ' very=True -> xlSheetVeryHidden
  On Error Resume Next
  If very Then Sh(sheetName).Visible = 2 Else Sh(sheetName).Visible = 0
  Chk "hide " & sheetName
End Sub

Sub SheetDelete(sheetName)
  On Error Resume Next
  Dim ws
  Set ws = Sh(sheetName)
  ws.Visible = -1
  ws.Delete
  Chk "delete " & sheetName
End Sub

Sub ClearAllComments()
  On Error Resume Next
  Dim ws, k
  For Each ws In wb.Worksheets
    For k = ws.CommentsThreaded.Count To 1 Step -1 : ws.CommentsThreaded(k).Delete : Next
    Err.Clear
    ws.Cells.ClearComments
  Next
  Chk "clear comments"
End Sub

Sub SetMeta(propName, value)            ' "" = clear
  On Error Resume Next
  wb.BuiltinDocumentProperties(propName).Value = U(value)
  If Err.Number <> 0 Then Note "metadata " & propName & " skipped: " & Err.Description
  Err.Clear
End Sub

Sub BreakAllLinks()                     ' only when the user asked for it explicitly
  On Error Resume Next
  Dim links, l
  links = wb.LinkSources(1)
  If IsArray(links) Then
    For Each l In links : wb.BreakLink l, 1 : Next
  End If
  Chk "break links"
End Sub

' New sheet that copies the layout of an existing sheet (the source sheet is only read)
Function NewSheetFromLayout(srcName, newName, copyRange, maxCol, maxRow, printArea)
  On Error Resume Next
  Dim s, n, w, i, z
  Set s = Sh(srcName)
  For Each w In wb.Worksheets
    If w.Name = U(newName) Then Fail "SHEET_EXISTS: " & newName
  Next
  If s.Index < wb.Sheets.Count Then
    Set n = wb.Worksheets.Add(wb.Sheets(s.Index + 1))
  Else
    Set n = wb.Worksheets.Add()
    n.Move , s
  End If
  Chk "add sheet"
  n.Name = U(newName) : Chk "rename sheet"
  s.Range(copyRange).Copy n.Range("A1") : Chk "copy layout"
  For i = 1 To maxCol : n.Columns(i).ColumnWidth = s.Columns(i).ColumnWidth : Next
  For i = 1 To maxRow : n.Rows(i).RowHeight = s.Rows(i).RowHeight : Next
  Chk "sizes"
  n.Tab.Color = s.Tab.Color : Err.Clear
  xl.PrintCommunication = False : Err.Clear
  n.PageSetup.Orientation = s.PageSetup.Orientation : Err.Clear
  n.PageSetup.PaperSize = s.PageSetup.PaperSize : Err.Clear
  n.PageSetup.LeftMargin = s.PageSetup.LeftMargin : Err.Clear
  n.PageSetup.RightMargin = s.PageSetup.RightMargin : Err.Clear
  n.PageSetup.TopMargin = s.PageSetup.TopMargin : Err.Clear
  n.PageSetup.BottomMargin = s.PageSetup.BottomMargin : Err.Clear
  n.PageSetup.HeaderMargin = s.PageSetup.HeaderMargin : Err.Clear
  n.PageSetup.FooterMargin = s.PageSetup.FooterMargin : Err.Clear
  n.PageSetup.CenterHorizontally = s.PageSetup.CenterHorizontally : Err.Clear
  z = s.PageSetup.Zoom
  If VarType(z) = vbBoolean Then
    n.PageSetup.Zoom = False : n.PageSetup.FitToPagesWide = s.PageSetup.FitToPagesWide : n.PageSetup.FitToPagesTall = s.PageSetup.FitToPagesTall
  Else
    n.PageSetup.Zoom = z
  End If
  Err.Clear
  If printArea <> "" Then n.PageSetup.PrintArea = printArea
  If Err.Number <> 0 Then Note "print area not set: " & Err.Description
  Err.Clear
  xl.PrintCommunication = True : Err.Clear
  ' gridlines like the source, cursor A1
  xl.Visible = True : xl.ScreenUpdating = False
  Dim g
  s.Activate : g = xl.ActiveWindow.DisplayGridlines
  n.Activate : xl.ActiveWindow.DisplayGridlines = g : xl.ActiveWindow.Zoom = 100 : n.Range("A1").Select
  xl.ScreenUpdating = True : xl.Visible = False
  Err.Clear
  Note "new sheet created: " & newName
  Set NewSheetFromLayout = n
End Function

' Plain text replace inside formulas of a range; stops everything if the count is not the expected one
Sub ReplaceInFormulas(sheetName, addr, findTxt, replTxt, expected)
  On Error Resume Next
  Dim c, f, f2, cnt
  cnt = 0
  For Each c In Sh(sheetName).Range(addr).Cells
    If c.HasFormula Then
      f = c.Formula
      f2 = Replace(f, U(findTxt), U(replTxt))
      If f2 <> f Then c.Formula = f2 : cnt = cnt + 1
    End If
  Next
  Chk "replace in " & sheetName & "!" & addr
  Note "replaced in " & cnt & " cells (" & sheetName & "!" & addr & ")"
  If expected >= 0 And cnt <> expected Then Fail "REPLACE_COUNT_MISMATCH expected " & expected & " got " & cnt
End Sub

Sub SetCellProtection(sheetName, addr, isLocked, isHidden)   ' sheet must be unprotected
  On Error Resume Next
  Dim r
  Set r = Sh(sheetName).Range(addr)
  r.Locked = isLocked : r.FormulaHidden = isHidden
  Chk "protection " & sheetName & "!" & addr & " (is the sheet protected?)"
End Sub

Sub LockFormulas(sheetName)             ' every formula cell -> Locked + Hidden (sheet must be unprotected)
  On Error Resume Next
  Dim r
  Set r = Sh(sheetName).UsedRange.SpecialCells(-4123)
  If Err.Number <> 0 Then Err.Clear : Note "no formulas in " & sheetName : Exit Sub
  r.Locked = True : r.FormulaHidden = True
  Chk "lock formulas " & sheetName & " (is the sheet protected?)"
  Note "formulas locked+hidden: " & sheetName & " " & r.Cells.Count & " cells"
End Sub

Sub TidyView()                          ' A1 on every visible sheet, open on the first visible sheet, ungroup
  On Error Resume Next
  Dim ws, first
  xl.Visible = True : xl.ScreenUpdating = False
  For Each ws In wb.Worksheets
    If ws.Visible = -1 Then
      If IsEmpty(first) Then Set first = ws
      ws.Select : xl.ActiveWindow.ScrollRow = 1 : xl.ActiveWindow.ScrollColumn = 1 : ws.Range("A1").Select
    End If
  Next
  If Not IsEmpty(first) Then first.Select
  xl.ScreenUpdating = True : xl.Visible = False
  Chk "tidy view"
  Note "tidy view done"
End Sub

Sub OpenOnSheet(sheetName)              ' optional: choose the sheet the file opens on
  On Error Resume Next
  xl.Visible = True : xl.ScreenUpdating = False
  Sh(sheetName).Select
  xl.ScreenUpdating = True : xl.Visible = False
  Chk "open on sheet"
End Sub

Sub RestoreActive()                     ' keep opening on the same sheet as the source file
  On Error Resume Next
  xl.Visible = True : xl.ScreenUpdating = False
  orig.Activate
  xl.ScreenUpdating = True : xl.Visible = False
  Err.Clear
End Sub

' ---------- main ----------
On Error Resume Next
If fso.FileExists(errFile) Then fso.DeleteFile errFile, True
Err.Clear
gCount = 0
If Not fso.FileExists(src) Then Fail "SOURCE_NOT_FOUND: " & src
If fso.FileExists(dst) Then Fail "TARGET_EXISTS: " & dst

Set xl = CreateObject("Excel.Application")
If Err.Number <> 0 Then Fail "EXCEL_NOT_AVAILABLE: " & Err.Description
xl.DisplayAlerts = False
xl.EnableEvents = False
xl.AutomationSecurity = 3               ' macros disabled while opening -> Workbook_Open does not run
Chk "EXCEL_SETUP"

If pwRequired Then
  ' Excel shows its OWN masked password box (the password never passes through chat or any file).
  ' A workbook opened with a password keeps that password when saved with SaveAs.
  xl.Visible = True
  Set wb = xl.Workbooks.Open(src, 0, True)
  xl.Visible = False
Else
  xl.Visible = False
  Set wb = xl.Workbooks.Open(src, 0, True)
End If
If Err.Number <> 0 Then Fail "OPEN_FAILED (wrong password, cancelled, or damaged file): " & Err.Description
Set orig = wb.ActiveSheet
Note "sheets before: " & wb.Worksheets.Count

'@@OPS@@

Note "sheets after: " & wb.Worksheets.Count
If LCase(Right(dst, 5)) = ".xlsm" Then wb.SaveAs dst, 52 Else wb.SaveAs dst, 51
If Err.Number <> 0 Then Fail "SAVE_FAILED: " & Err.Description
Note "SAVED: " & dst
Finish 0
