Attribute VB_Name = "TeklaClashClassifier"
Option Explicit

' Tekla Clash Classifier
' One entry point: RunClashClassification (assigned to the RUN button)
' Flow: pick Universal -> Bolt -> ClashCheck.log -> write NEW workbook with 4 tabs.
' The classifier workbook itself is never overwritten with result sheets.

Private Const SHEET_CONTROL As String = "Control"
Private Const SHEET_PART_PART As String = "Part_to_Part_Clash"
Private Const SHEET_PART_BOLT As String = "Part_to_Bolt_Clash"
Private Const SHEET_BOLT_BOLT As String = "Bolt_to_Bolt_Clash"
Private Const SHEET_DUP As String = "Duplicate members"
Private Const SHEET_SUMMARY As String = "Summary"
Private Const SHEET_UNRESOLVED As String = "Unresolved"
Private Const RUN_BUTTON_NAME As String = "btnRunClashClassification"

' Cached regex objects (creating RegExp per cell was a major slowdown)
Private mGuidRe As Object
Private mClashRe As Object

' Runs when the macro-enabled workbook opens (after VBA is imported).
Public Sub Auto_Open()
    On Error Resume Next
    SetupClassifierUI
End Sub

' One-time / repair: rebuild Control sheet + RUN Form Control button.
Public Sub SetupClassifierUI()
    Dim ws As Worksheet
    EnsureControlSheet
    Set ws = ThisWorkbook.Sheets(SHEET_CONTROL)
    EnsureRunButton ws
End Sub

' === MAIN ENTRY (wire this to the RUN button) ===
Public Sub RunClashClassification()
    Dim universalPath As String
    Dim boltPath As String
    Dim clashPath As String
    Dim outPath As Variant
    Dim parts As Object
    Dim bolts As Object
    Dim clashes As Collection
    Dim counts As Object
    Dim outWb As Workbook
    Dim statusLine As String

    On Error GoTo Fail

    EnsureControlSheet
    EnsureRunButton ThisWorkbook.Sheets(SHEET_CONTROL)

    ' 1) Universal Report
    universalPath = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", _
        Title:="Step 1 of 3 - Select Universal Report", _
        MultiSelect:=False)
    If universalPath = "False" Or Len(universalPath) = 0 Then
        MsgBox "Cancelled: Universal Report was not selected.", vbExclamation, "Tekla Clash Classifier"
        Exit Sub
    End If

    ' 2) Bolt Checking Report
    boltPath = Application.GetOpenFilename( _
        FileFilter:="Bolt Report (*.xlsx;*.xlsm;*.xls;*.csv),*.xlsx;*.xlsm;*.xls;*.csv", _
        Title:="Step 2 of 3 - Select Bolt Checking Report", _
        MultiSelect:=False)
    If boltPath = "False" Or Len(boltPath) = 0 Then
        MsgBox "Cancelled: Bolt Checking Report was not selected.", vbExclamation, "Tekla Clash Classifier"
        Exit Sub
    End If

    ' 3) ClashCheck.log
    clashPath = Application.GetOpenFilename( _
        FileFilter:="Clash Log (*.log;*.txt),*.log;*.txt,All Files (*.*),*.*", _
        Title:="Step 3 of 3 - Select ClashCheck.log", _
        MultiSelect:=False)
    If clashPath = "False" Or Len(clashPath) = 0 Then
        MsgBox "Cancelled: ClashCheck.log was not selected.", vbExclamation, "Tekla Clash Classifier"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error Resume Next
    Application.PrintCommunication = False
    On Error GoTo Fail

    Application.StatusBar = "Loading Universal Report..."
    Set parts = LoadParts(CStr(universalPath))
    Application.StatusBar = "Loading Bolt report..."
    Set bolts = LoadBolts(CStr(boltPath))
    Application.StatusBar = "Parsing ClashCheck.log..."
    Set clashes = ParseClashLog(CStr(clashPath))

    Application.StatusBar = "Classifying clashes (bulk write)..."
    Set outWb = Workbooks.Add(xlWBATWorksheet)
    Set counts = ClassifyAndWrite(outWb, clashes, parts, bolts)
    WriteSummary outWb, CStr(universalPath), CStr(boltPath), CStr(clashPath), _
        parts.Count, bolts.Count, clashes.Count, counts

    ' Put Summary first, then the four category tabs
    On Error Resume Next
    outWb.Sheets(SHEET_SUMMARY).Move Before:=outWb.Sheets(1)
    On Error GoTo Fail

    Application.StatusBar = False
    Application.EnableEvents = True
    On Error Resume Next
    Application.PrintCommunication = True
    On Error GoTo Fail
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    outPath = Application.GetSaveAsFilename( _
        InitialFileName:="Tekla_Clash_Classification.xlsx", _
        FileFilter:="Excel Workbook (*.xlsx),*.xlsx", _
        Title:="Save classified clash results")
    If outPath = False Or Len(CStr(outPath)) = 0 Then
        ' Leave the new workbook open unsaved; classifier XLSM untouched.
        statusLine = "Complete (output not saved yet) | " & CountSummary(counts)
        UpdateControlStatus "OK (unsaved output)", statusLine
        MsgBox "Classification complete. Save the new workbook when ready." & vbCrLf & statusLine, _
            vbInformation, "Tekla Clash Classifier"
        Exit Sub
    End If

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    outWb.SaveAs Filename:=CStr(outPath), FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    statusLine = CountSummary(counts)
    UpdateControlStatus "OK " & Format$(Now, "yyyy-mm-dd hh:nn:ss"), statusLine & " | Saved: " & CStr(outPath)

    MsgBox "Clash classification complete." & vbCrLf & statusLine & vbCrLf & "Saved: " & CStr(outPath), _
        vbInformation, "Tekla Clash Classifier"
    Exit Sub

Fail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False
    Application.DisplayAlerts = True
    On Error Resume Next
    Application.PrintCommunication = True
    On Error GoTo 0
    UpdateControlStatus "ERROR", Err.Description
    MsgBox "Error: " & Err.Description, vbCritical, "Tekla Clash Classifier"
End Sub

Private Function CountSummary(ByVal counts As Object) As String
    CountSummary = _
        "Part-Part=" & counts(SHEET_PART_PART) & _
        " | Part-Bolt=" & counts(SHEET_PART_BOLT) & _
        " | Bolt-Bolt=" & counts(SHEET_BOLT_BOLT) & _
        " | Duplicate=" & counts(SHEET_DUP) & _
        " | Unresolved=" & counts(SHEET_UNRESOLVED)
End Function

Private Sub UpdateControlStatus(ByVal statusText As String, ByVal countsText As String)
    On Error Resume Next
    With ThisWorkbook.Sheets(SHEET_CONTROL)
        .Range("B5").Value = statusText
        .Range("B6").Value = countsText
    End With
End Sub

Public Sub EnsureControlSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_CONTROL)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = SHEET_CONTROL
    End If

    ws.Cells.Clear
    ' Remove leftover Form/ActiveX buttons except we recreate RUN below
    DeleteAllButtons ws

    ws.Range("A1").Value = "Tekla Clash Classifier"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Color = RGB(255, 255, 255)
    ws.Range("A1").Interior.Color = RGB(31, 78, 121)
    ws.Range("A1:D1").Merge
    ws.Rows(1).RowHeight = 28

    ws.Range("A2").Value = "Click RUN. You will be asked for 3 files one by one (Universal Report -> Bolt Report -> ClashCheck.log). Results are saved to a NEW Excel workbook - this classifier file stays unchanged."
    ws.Range("A2").WrapText = True
    ws.Range("A2:D2").Merge
    ws.Rows(2).RowHeight = 48

    ws.Range("A3").Value = "Output tabs"
    ws.Range("A3").Font.Bold = True
    ws.Range("B3").Value = "Part_to_Part_Clash | Part_to_Bolt_Clash | Bolt_to_Bolt_Clash | Duplicate members"

    ws.Range("A5").Value = "Last status"
    ws.Range("A5").Font.Bold = True
    ws.Range("A6").Value = "Last counts"
    ws.Range("A6").Font.Bold = True
    ws.Range("B5").Value = ""
    ws.Range("B6").Value = ""

    ws.Range("A11").Value = "If macros blocked"
    ws.Range("A11").Font.Bold = True
    ws.Range("B11").Value = "Close Excel. Right-click this .xlsm in File Explorer -> Properties -> check Unblock -> OK. Re-open and Enable Content. Then import TeklaClashClassifier.bas if not already imported."
    ws.Range("B11").WrapText = True
    ws.Rows(11).RowHeight = 40

    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B").ColumnWidth = 90
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 14
End Sub

Private Sub EnsureRunButton(ByVal ws As Worksheet)
    Dim btn As Button
    Dim leftPos As Double, topPos As Double

    On Error Resume Next
    ws.Buttons(RUN_BUTTON_NAME).Delete
    On Error GoTo 0

    ' Place RUN button under the instructions (row 8 area)
    leftPos = ws.Range("A8").Left
    topPos = ws.Range("A8").Top
    Set btn = ws.Buttons.Add(leftPos, topPos, 120, 36)
    btn.Name = RUN_BUTTON_NAME
    btn.Caption = "RUN"
    ' Prefer workbook-local macro reference (matches embedded VML FmlaMacro).
    btn.OnAction = "RunClashClassification"
    On Error Resume Next
    btn.OnAction = "'" & ThisWorkbook.Name & "'!RunClashClassification"
    On Error GoTo 0
    btn.Font.Bold = True
    btn.Font.Size = 14
End Sub

Private Sub DeleteAllButtons(ByVal ws As Worksheet)
    Dim b As Button
    On Error Resume Next
    For Each b In ws.Buttons
        b.Delete
    Next b
    On Error GoTo 0
End Sub

Private Function GuidRegex() As Object
    If mGuidRe Is Nothing Then
        Set mGuidRe = CreateObject("VBScript.RegExp")
        mGuidRe.Global = False
        mGuidRe.IgnoreCase = True
        mGuidRe.Pattern = "([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})"
    End If
    Set GuidRegex = mGuidRe
End Function

Private Function ClashRegex() As Object
    If mClashRe Is Nothing Then
        Set mClashRe = CreateObject("VBScript.RegExp")
        mClashRe.Global = False
        mClashRe.IgnoreCase = True
        mClashRe.Pattern = "Guid:\s*([0-9a-fA-F-]{36})\s*<->\s*Guid:\s*([0-9a-fA-F-]{36})"
    End If
    Set ClashRegex = mClashRe
End Function

Private Function NormalizeGuid(ByVal value As Variant) As String
    Dim s As String
    Dim m As Object
    s = Trim$(CStr(value & ""))
    If UCase$(Left$(s, 5)) = "GUID:" Then s = Trim$(Mid$(s, 6))
    s = Trim$(s)
    ' Fast path: already a bare GUID
    If Len(s) = 36 And Mid$(s, 9, 1) = "-" Then
        NormalizeGuid = LCase$(s)
        Exit Function
    End If
    If GuidRegex().Test(s) Then
        Set m = GuidRegex().Execute(s)
        NormalizeGuid = LCase$(m(0).SubMatches(0))
    Else
        NormalizeGuid = LCase$(s)
    End If
End Function

Private Function FormatGuid(ByVal guid As String) As String
    If Len(guid) = 0 Then
        FormatGuid = ""
    Else
        FormatGuid = "GUID:" & guid
    End If
End Function

Private Function LoadParts(ByVal path As String) As Object
    Dim dict As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim headerRow As Long, guidCol As Long
    Dim profileCol As Long, nameCol As Long
    Dim lastRow As Long, lastCol As Long, r As Long
    Dim guid As String
    Dim data As Variant
    Dim profile As String, nm As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1 ' TextCompare

    Set wb = Workbooks.Open(Filename:=path, ReadOnly:=True, UpdateLinks:=0)
    Set ws = PickBestSheet(wb, "universal,report,part")
    FindHeaderRowAndGuidCol ws, headerRow, guidCol
    profileCol = FindColByNames(ws, headerRow, Array("Profile"))
    nameCol = FindColByNames(ws, headerRow, Array("Name"))
    lastRow = ws.Cells(ws.Rows.Count, guidCol).End(xlUp).Row
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < guidCol Then lastCol = guidCol
    If profileCol > lastCol Then lastCol = profileCol
    If nameCol > lastCol Then lastCol = nameCol

    ' One bulk read instead of per-cell COM calls
    data = ws.Range(ws.Cells(headerRow + 1, 1), ws.Cells(lastRow, lastCol)).Value
    If IsArray(data) Then
        For r = 1 To UBound(data, 1)
            guid = NormalizeGuid(data(r, guidCol))
            If InStr(guid, "-") > 0 Then
                profile = ""
                nm = ""
                If profileCol > 0 Then profile = Trim$(CStr(data(r, profileCol) & ""))
                If nameCol > 0 Then nm = Trim$(CStr(data(r, nameCol) & ""))
                If Not dict.Exists(guid) Then dict.Add guid, Array(profile, nm)
            End If
        Next r
    End If

    wb.Close SaveChanges:=False
    Set LoadParts = dict
End Function

Private Function LoadBolts(ByVal path As String) As Object
    Dim dict As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim headerRow As Long, guidCol As Long
    Dim diaCol As Long, stdCol As Long
    Dim lastRow As Long, lastCol As Long, r As Long
    Dim guid As String
    Dim diaVal As Variant
    Dim data As Variant
    Dim std As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    Set wb = Workbooks.Open(Filename:=path, ReadOnly:=True, Local:=True, UpdateLinks:=0)
    Set ws = PickBestSheet(wb, "bolt,report")
    FindHeaderRowAndGuidCol ws, headerRow, guidCol
    diaCol = FindColByNames(ws, headerRow, Array("BOLT_DIA", "Bolt_Dia", "Diameter", "DIA"))
    stdCol = FindColByNames(ws, headerRow, Array("BOLT_STD", "Bolt_Std", "Standard"))
    lastRow = ws.Cells(ws.Rows.Count, guidCol).End(xlUp).Row
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < guidCol Then lastCol = guidCol
    If diaCol > lastCol Then lastCol = diaCol
    If stdCol > lastCol Then lastCol = stdCol

    data = ws.Range(ws.Cells(headerRow + 1, 1), ws.Cells(lastRow, lastCol)).Value
    If IsArray(data) Then
        For r = 1 To UBound(data, 1)
            guid = NormalizeGuid(data(r, guidCol))
            If InStr(guid, "-") > 0 Then
                diaVal = ""
                If diaCol > 0 Then
                    diaVal = data(r, diaCol)
                    If VarType(diaVal) = vbString Then
                        diaVal = Trim$(CStr(diaVal))
                        If IsNumeric(diaVal) Then diaVal = CLng(diaVal)
                    End If
                End If
                std = ""
                If stdCol > 0 Then std = Trim$(CStr(data(r, stdCol) & ""))
                If Not dict.Exists(guid) Then dict.Add guid, Array(diaVal, std)
            End If
        Next r
    End If

    wb.Close SaveChanges:=False
    Set LoadBolts = dict
End Function

Private Function ParseClashLog(ByVal path As String) As Collection
    Dim col As Collection
    Dim fnum As Integer
    Dim line As String
    Dim re As Object
    Dim m As Object
    Dim g1 As String, g2 As String
    Dim isInside As Boolean

    Set col = New Collection
    Set re = ClashRegex()

    fnum = FreeFile
    Open path For Input As #fnum
    Do While Not EOF(fnum)
        Line Input #fnum, line
        If re.Test(line) Then
            Set m = re.Execute(line)
            g1 = LCase$(m(0).SubMatches(0))
            g2 = LCase$(m(0).SubMatches(1))
            isInside = (InStr(1, line, "inside member", vbTextCompare) > 0)
            col.Add Array(g1, g2, isInside, line)
        End If
    Loop
    Close #fnum
    Set ParseClashLog = col
End Function

Private Function KindOf(ByVal guid As String, ByVal parts As Object, ByVal bolts As Object) As String
    Dim inPart As Boolean, inBolt As Boolean
    inPart = parts.Exists(guid)
    inBolt = bolts.Exists(guid)
    If inPart And Not inBolt Then
        KindOf = "part"
    ElseIf inBolt And Not inPart Then
        KindOf = "bolt"
    ElseIf inPart And inBolt Then
        KindOf = "both"
    Else
        KindOf = "unknown"
    End If
End Function

Private Function AttrFor(ByVal guid As String, ByVal parts As Object, ByVal bolts As Object) As Variant
    If parts.Exists(guid) Then
        AttrFor = parts(guid)(0)
    ElseIf bolts.Exists(guid) Then
        AttrFor = bolts(guid)(0)
    Else
        AttrFor = ""
    End If
End Function

Private Function ClassifyAndWrite(ByVal outWb As Workbook, ByVal clashes As Collection, _
    ByVal parts As Object, ByVal bolts As Object) As Object

    Dim counts As Object
    Dim i As Long, maxN As Long
    Dim item As Variant
    Dim g1 As String, g2 As String
    Dim k1 As String, k2 As String
    Dim isInside As Boolean
    Dim wsPP As Worksheet, wsPB As Worksheet, wsBB As Worksheet, wsDup As Worksheet, wsU As Worksheet
    Dim nPP As Long, nPB As Long, nBB As Long, nDup As Long, nU As Long
    Dim arrPP() As Variant, arrPB() As Variant, arrBB() As Variant, arrDup() As Variant, arrU() As Variant
    Dim tg As String, tk As String

    Set counts = CreateObject("Scripting.Dictionary")
    counts.Add SHEET_PART_PART, 0
    counts.Add SHEET_PART_BOLT, 0
    counts.Add SHEET_BOLT_BOLT, 0
    counts.Add SHEET_DUP, 0
    counts.Add SHEET_UNRESOLVED, 0

    Set wsPP = AddOutputSheet(outWb, SHEET_PART_PART, "PROFILE1", "PROFILE2")
    Set wsPB = AddOutputSheet(outWb, SHEET_PART_BOLT, "PROFILE1", "bolt1")
    Set wsBB = AddOutputSheet(outWb, SHEET_BOLT_BOLT, "bolt1", "bolt2")
    Set wsDup = AddOutputSheet(outWb, SHEET_DUP, "PROFILE1", "PROFILE2")

    maxN = clashes.Count
    If maxN < 1 Then maxN = 1
    ReDim arrPP(1 To maxN, 1 To 5)
    ReDim arrPB(1 To maxN, 1 To 5)
    ReDim arrBB(1 To maxN, 1 To 5)
    ReDim arrDup(1 To maxN, 1 To 5)
    ReDim arrU(1 To maxN, 1 To 5)
    nPP = 0: nPB = 0: nBB = 0: nDup = 0: nU = 0
    Set wsU = Nothing

    For i = 1 To clashes.Count
        item = clashes(i)
        g1 = item(0)
        g2 = item(1)
        isInside = item(2)
        k1 = KindOf(g1, parts, bolts)
        k2 = KindOf(g2, parts, bolts)

        If isInside Then
            nDup = nDup + 1
            arrDup(nDup, 1) = FormatGuid(g1)
            arrDup(nDup, 2) = FormatGuid(g2)
            arrDup(nDup, 3) = AttrFor(g1, parts, bolts)
            arrDup(nDup, 4) = AttrFor(g2, parts, bolts)
            arrDup(nDup, 5) = ""
        Else
            If k1 = "bolt" And k2 = "part" Then
                tg = g1: g1 = g2: g2 = tg
                tk = k1: k1 = k2: k2 = tk
            End If

            If k1 = "part" And k2 = "part" Then
                nPP = nPP + 1
                arrPP(nPP, 1) = FormatGuid(g1)
                arrPP(nPP, 2) = FormatGuid(g2)
                arrPP(nPP, 3) = AttrFor(g1, parts, bolts)
                arrPP(nPP, 4) = AttrFor(g2, parts, bolts)
                arrPP(nPP, 5) = ""
            ElseIf k1 = "part" And k2 = "bolt" Then
                nPB = nPB + 1
                arrPB(nPB, 1) = FormatGuid(g1)
                arrPB(nPB, 2) = FormatGuid(g2)
                arrPB(nPB, 3) = AttrFor(g1, parts, bolts)
                arrPB(nPB, 4) = AttrFor(g2, parts, bolts)
                arrPB(nPB, 5) = ""
            ElseIf k1 = "bolt" And k2 = "bolt" Then
                nBB = nBB + 1
                arrBB(nBB, 1) = FormatGuid(g1)
                arrBB(nBB, 2) = FormatGuid(g2)
                arrBB(nBB, 3) = AttrFor(g1, parts, bolts)
                arrBB(nBB, 4) = AttrFor(g2, parts, bolts)
                arrBB(nBB, 5) = ""
            Else
                If wsU Is Nothing Then
                    Set wsU = outWb.Worksheets.Add(After:=outWb.Sheets(outWb.Sheets.Count))
                    wsU.Name = SHEET_UNRESOLVED
                    wsU.Range("A1").Value = SHEET_UNRESOLVED
                    wsU.Range("A1").Font.Bold = True
                    wsU.Range("A1").Font.Size = 14
                    wsU.Range("A1:E1").Merge
                    wsU.Range("A2").Value = "GUID1"
                    wsU.Range("B2").Value = "GUID2"
                    wsU.Range("C2").Value = "Kind1"
                    wsU.Range("D2").Value = "Kind2"
                    wsU.Range("E2").Value = "Log line"
                    wsU.Range("A2:E2").Font.Bold = True
                End If
                nU = nU + 1
                arrU(nU, 1) = FormatGuid(g1)
                arrU(nU, 2) = FormatGuid(g2)
                arrU(nU, 3) = k1
                arrU(nU, 4) = k2
                arrU(nU, 5) = item(3)
            End If
        End If
    Next i

    ' One Range dump per sheet (much faster than writing each cell)
    DumpRows wsPP, arrPP, nPP
    DumpRows wsPB, arrPB, nPB
    DumpRows wsBB, arrBB, nBB
    DumpRows wsDup, arrDup, nDup
    If nU > 0 And Not wsU Is Nothing Then DumpRows wsU, arrU, nU

    counts(SHEET_PART_PART) = nPP
    counts(SHEET_PART_BOLT) = nPB
    counts(SHEET_BOLT_BOLT) = nBB
    counts(SHEET_DUP) = nDup
    counts(SHEET_UNRESOLVED) = nU

    SortSheetByAttrColumns wsPP
    SortSheetByAttrColumns wsPB
    SortSheetByAttrColumns wsBB
    SortSheetByAttrColumns wsDup

    CleanupDefaultSheets outWb

    Set ClassifyAndWrite = counts
End Function

Private Sub DumpRows(ByVal ws As Worksheet, ByRef src As Variant, ByVal n As Long)
    Dim out() As Variant
    Dim i As Long, j As Long
    If n < 1 Then Exit Sub
    ReDim out(1 To n, 1 To 5)
    For i = 1 To n
        For j = 1 To 5
            out(i, j) = src(i, j)
        Next j
    Next i
    ws.Range("A3").Resize(n, 5).Value = out
End Sub

Private Sub SortSheetByAttrColumns(ByVal ws As Worksheet)
    ' Sort full data block A3:E(last) by Col C then Col D (PROFILE1/bolt1, PROFILE2/bolt2).
    ' Entire rows move together so GUID1/GUID2/Remarks stay aligned.
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 3 Then Exit Sub

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("C3:C" & lastRow), Order:=xlAscending
        .SortFields.Add Key:=ws.Range("D3:D" & lastRow), Order:=xlAscending
        .SetRange ws.Range("A3:E" & lastRow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

Private Function AddOutputSheet(ByVal outWb As Workbook, ByVal sheetName As String, _
    ByVal col3Header As String, ByVal col4Header As String) As Worksheet
    Dim ws As Worksheet
    Set ws = outWb.Worksheets.Add(After:=outWb.Sheets(outWb.Sheets.Count))
    ws.Name = sheetName
    ' Row 1: title
    ws.Range("A1").Value = sheetName
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A1:E1").Merge
    ' Row 2: column headers
    ws.Range("A2").Value = "GUID1"
    ws.Range("B2").Value = "GUID2"
    ws.Range("C2").Value = col3Header
    ws.Range("D2").Value = col4Header
    ws.Range("E2").Value = "Remarks"
    ws.Range("A2:E2").Font.Bold = True
    ws.Columns("A").ColumnWidth = 42
    ws.Columns("B").ColumnWidth = 42
    ws.Columns("C").ColumnWidth = 22
    ws.Columns("D").ColumnWidth = 22
    ws.Columns("E").ColumnWidth = 40
    Set AddOutputSheet = ws
End Function

Private Sub CleanupDefaultSheets(ByVal outWb As Workbook)
    Dim ws As Worksheet
    Dim nm As String
    On Error Resume Next
    Application.DisplayAlerts = False
    For Each ws In outWb.Worksheets
        nm = ws.Name
        If nm = "Sheet1" Or nm = "Sheet2" Or nm = "Sheet3" Then
            If outWb.Worksheets.Count > 1 Then ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Private Sub WriteSummary(ByVal outWb As Workbook, ByVal u As String, ByVal b As String, ByVal c As String, _
    ByVal partCount As Long, ByVal boltCount As Long, ByVal clashCount As Long, ByVal counts As Object)

    Dim ws As Worksheet
    Set ws = outWb.Worksheets.Add(After:=outWb.Sheets(outWb.Sheets.Count))
    ws.Name = SHEET_SUMMARY
    ws.Range("A1").Value = "Tekla Clash Classification Summary"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A3").Value = "Universal Report": ws.Range("B3").Value = u
    ws.Range("A4").Value = "Bolt Report": ws.Range("B4").Value = b
    ws.Range("A5").Value = "ClashCheck.log": ws.Range("B5").Value = c
    ws.Range("A7").Value = "Parts loaded": ws.Range("B7").Value = partCount
    ws.Range("A8").Value = "Bolts loaded": ws.Range("B8").Value = boltCount
    ws.Range("A9").Value = "Clash pairs parsed": ws.Range("B9").Value = clashCount
    ws.Range("A11").Value = "Sheet": ws.Range("B11").Value = "Rows"
    ws.Range("A11:B11").Font.Bold = True
    ws.Range("A12").Value = SHEET_PART_PART: ws.Range("B12").Value = counts(SHEET_PART_PART)
    ws.Range("A13").Value = SHEET_PART_BOLT: ws.Range("B13").Value = counts(SHEET_PART_BOLT)
    ws.Range("A14").Value = SHEET_BOLT_BOLT: ws.Range("B14").Value = counts(SHEET_BOLT_BOLT)
    ws.Range("A15").Value = SHEET_DUP: ws.Range("B15").Value = counts(SHEET_DUP)
    ws.Range("A16").Value = SHEET_UNRESOLVED: ws.Range("B16").Value = counts(SHEET_UNRESOLVED)
    ws.Columns("A").ColumnWidth = 28
    ws.Columns("B").ColumnWidth = 80
End Sub

Private Function PickBestSheet(ByVal wb As Workbook, ByVal keywordsCsv As String) As Worksheet
    Dim ws As Worksheet
    Dim keys() As String
    Dim i As Long
    Dim best As Worksheet
    Dim bestScore As Long, score As Long
    Dim nm As String

    keys = Split(LCase$(keywordsCsv), ",")
    Set best = wb.Sheets(1)
    bestScore = -1
    For Each ws In wb.Worksheets
        nm = LCase$(ws.Name)
        score = 0
        For i = LBound(keys) To UBound(keys)
            If InStr(nm, Trim$(keys(i))) > 0 Then score = score + 1
        Next i
        If score > bestScore Then
            bestScore = score
            Set best = ws
        End If
    Next ws
    Set PickBestSheet = best
End Function

Private Sub FindHeaderRowAndGuidCol(ByVal ws As Worksheet, ByRef headerRow As Long, ByRef guidCol As Long)
    Dim r As Long, c As Long
    Dim lastCol As Long
    Dim v As String
    Dim compact As String
    Dim bestScore As Long, score As Long
    Dim bestRow As Long, bestCol As Long
    Dim hasProfile As Boolean, hasBoltDia As Boolean

    bestScore = -1
    bestRow = 1
    bestCol = 1

    For r = 1 To 40
        lastCol = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        If lastCol < 1 Then GoTo NextR
        score = 0
        hasProfile = False
        hasBoltDia = False
        Dim gCol As Long
        gCol = 0
        For c = 1 To lastCol
            v = Trim$(CStr(ws.Cells(r, c).Value & ""))
            If Len(v) = 0 Then GoTo NextC
            compact = LCase$(Replace(v, " ", ""))
            If compact = "guid" Then
                gCol = c
                score = score + 10
            ElseIf Left$(compact, 4) = "guid" And gCol = 0 Then
                gCol = c
                score = score + 5
            End If
            If compact = "profile" Then hasProfile = True
            If InStr(compact, "bolt_dia") > 0 Or compact = "boltdia" Or compact = "diameter" Then hasBoltDia = True
NextC:
        Next c
        If gCol > 0 Then
            If hasProfile Then score = score + 5
            If hasBoltDia Then score = score + 5
            If score > bestScore Then
                bestScore = score
                bestRow = r
                bestCol = gCol
            End If
        End If
NextR:
    Next r

    If bestScore < 0 Then
        Err.Raise vbObjectError + 10, , "Could not find GUID column/header in report: " & ws.Parent.Name & " / " & ws.Name
    End If
    headerRow = bestRow
    guidCol = bestCol
End Sub

Private Function FindColByNames(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal names As Variant) As Long
    Dim c As Long, lastCol As Long
    Dim v As String, compact As String
    Dim i As Long, want As String

    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        v = Trim$(CStr(ws.Cells(headerRow, c).Value & ""))
        compact = LCase$(Replace(v, " ", ""))
        For i = LBound(names) To UBound(names)
            want = LCase$(Replace(CStr(names(i)), " ", ""))
            If compact = want Or InStr(compact, want) > 0 Then
                FindColByNames = c
                Exit Function
            End If
        Next i
    Next c
    FindColByNames = 0
End Function
