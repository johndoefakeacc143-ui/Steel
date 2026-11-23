Attribute VB_Name = "TeklaClashClassifier"
Option Explicit

' Tekla Clash Classifier
' Entry point: RunClashClassification
' Sheets written:
'   Part_to_Part_Clash
'   Part_to_Bolt_Clash
'   Bolt_to_Bolt_Clash
'   Duplicate members
' Plus Summary / Unresolved

Private Const SHEET_CONTROL As String = "Control"
Private Const SHEET_PART_PART As String = "Part_to_Part_Clash"
Private Const SHEET_PART_BOLT As String = "Part_to_Bolt_Clash"
Private Const SHEET_BOLT_BOLT As String = "Bolt_to_Bolt_Clash"
Private Const SHEET_DUP As String = "Duplicate members"
Private Const SHEET_SUMMARY As String = "Summary"
Private Const SHEET_UNRESOLVED As String = "Unresolved"

Public Sub RunClashClassification()
    Dim universalPath As String
    Dim boltPath As String
    Dim clashPath As String
    Dim parts As Object
    Dim bolts As Object
    Dim clashes As Collection
    Dim counts As Object

    On Error GoTo Fail

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    EnsureControlSheet
    universalPath = Trim$(CStr(ThisWorkbook.Sheets(SHEET_CONTROL).Range("B3").Value))
    boltPath = Trim$(CStr(ThisWorkbook.Sheets(SHEET_CONTROL).Range("B4").Value))
    clashPath = Trim$(CStr(ThisWorkbook.Sheets(SHEET_CONTROL).Range("B5").Value))

    If Len(universalPath) = 0 Or Len(Dir$(universalPath)) = 0 Then
        Err.Raise vbObjectError + 1, , "Universal Report path is missing or file not found. Use Browse or fill Control!B3."
    End If
    If Len(boltPath) = 0 Or Len(Dir$(boltPath)) = 0 Then
        Err.Raise vbObjectError + 2, , "Bolt report path is missing or file not found. Use Browse or fill Control!B4."
    End If
    If Len(clashPath) = 0 Or Len(Dir$(clashPath)) = 0 Then
        Err.Raise vbObjectError + 3, , "ClashCheck.log path is missing or file not found. Use Browse or fill Control!B5."
    End If

    Set parts = LoadParts(universalPath)
    Set bolts = LoadBolts(boltPath)
    Set clashes = ParseClashLog(clashPath)
    Set counts = ClassifyAndWrite(clashes, parts, bolts)

    WriteSummary universalPath, boltPath, clashPath, parts.Count, bolts.Count, clashes.Count, counts

    ThisWorkbook.Sheets(SHEET_CONTROL).Range("B7").Value = "OK " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ThisWorkbook.Sheets(SHEET_CONTROL).Range("B8").Value = _
        "Part-Part=" & counts(SHEET_PART_PART) & _
        " | Part-Bolt=" & counts(SHEET_PART_BOLT) & _
        " | Bolt-Bolt=" & counts(SHEET_BOLT_BOLT) & _
        " | Duplicate=" & counts(SHEET_DUP) & _
        " | Unresolved=" & counts(SHEET_UNRESOLVED)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Clash classification complete." & vbCrLf & ThisWorkbook.Sheets(SHEET_CONTROL).Range("B8").Value, vbInformation, "Tekla Clash Classifier"
    Exit Sub

Fail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ThisWorkbook.Sheets(SHEET_CONTROL).Range("B7").Value = "ERROR"
    ThisWorkbook.Sheets(SHEET_CONTROL).Range("B8").Value = Err.Description
    MsgBox "Error: " & Err.Description, vbCritical, "Tekla Clash Classifier"
End Sub

Public Sub BrowseUniversal()
    Dim p As String
    p = PickFile("Select Universal Report (*.xlsx;*.xlsm;*.csv)", "Excel/CSV Files,*.xlsx;*.xlsm;*.xls;*.csv")
    If Len(p) > 0 Then ThisWorkbook.Sheets(SHEET_CONTROL).Range("B3").Value = p
End Sub

Public Sub BrowseBolt()
    Dim p As String
    p = PickFile("Select Bolt Checking Report (*.xlsx;*.csv)", "Excel/CSV Files,*.xlsx;*.xlsm;*.xls;*.csv")
    If Len(p) > 0 Then ThisWorkbook.Sheets(SHEET_CONTROL).Range("B4").Value = p
End Sub

Public Sub BrowseClashLog()
    Dim p As String
    p = PickFile("Select ClashCheck.log", "Log Files,*.log;*.txt,All Files,*.*")
    If Len(p) > 0 Then ThisWorkbook.Sheets(SHEET_CONTROL).Range("B5").Value = p
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
    ws.Range("A1").Value = "Tekla Clash Classifier"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A2").Value = "Set the three input paths (or use Browse macros), then run RunClashClassification."
    ws.Range("A3").Value = "Universal Report"
    ws.Range("A4").Value = "Bolt Report"
    ws.Range("A5").Value = "ClashCheck.log"
    ws.Range("A7").Value = "Last status"
    ws.Range("A8").Value = "Last counts"
    ws.Columns("A").ColumnWidth = 22
    ws.Columns("B").ColumnWidth = 90
End Sub

Private Function PickFile(ByVal title As String, ByVal filter As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .Title = title
        .Filters.Clear
        ' filter string like "Excel,*.xlsx;*.csv"
        Dim parts() As String
        parts = Split(filter, ",")
        Dim i As Long
        For i = 0 To UBound(parts) Step 2
            If i + 1 <= UBound(parts) Then
                .Filters.Add parts(i), parts(i + 1)
            End If
        Next i
        If .Show = -1 Then
            PickFile = .SelectedItems(1)
        Else
            PickFile = ""
        End If
    End With
End Function

Private Function NormalizeGuid(ByVal value As Variant) As String
    Dim s As String
    Dim re As Object
    Dim m As Object
    s = Trim$(CStr(value & ""))
    If UCase$(Left$(s, 5)) = "GUID:" Then s = Trim$(Mid$(s, 6))
    s = Trim$(s)
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})"
    If re.Test(s) Then
        Set m = re.Execute(s)
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
    Dim lastRow As Long, r As Long
    Dim guid As String
    Dim opened As Boolean

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1 ' TextCompare

    If LCase$(Right$(path, 4)) = ".csv" Then
        Set wb = Workbooks.Open(Filename:=path, ReadOnly:=True, Local:=True)
        opened = True
    Else
        Set wb = Workbooks.Open(Filename:=path, ReadOnly:=True)
        opened = True
    End If

    Set ws = PickBestSheet(wb, "universal,report,part")
    FindHeaderRowAndGuidCol ws, headerRow, guidCol
    profileCol = FindColByNames(ws, headerRow, Array("Profile"))
    nameCol = FindColByNames(ws, headerRow, Array("Name"))
    lastRow = ws.Cells(ws.Rows.Count, guidCol).End(xlUp).Row

    For r = headerRow + 1 To lastRow
        guid = NormalizeGuid(ws.Cells(r, guidCol).Value)
        If InStr(guid, "-") > 0 Then
            Dim profile As String, nm As String
            profile = ""
            nm = ""
            If profileCol > 0 Then profile = Trim$(CStr(ws.Cells(r, profileCol).Value & ""))
            If nameCol > 0 Then nm = Trim$(CStr(ws.Cells(r, nameCol).Value & ""))
            If Not dict.Exists(guid) Then
                dict.Add guid, Array(profile, nm)
            End If
        End If
    Next r

    If opened Then wb.Close SaveChanges:=False
    Set LoadParts = dict
End Function

Private Function LoadBolts(ByVal path As String) As Object
    Dim dict As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim headerRow As Long, guidCol As Long
    Dim diaCol As Long, stdCol As Long
    Dim lastRow As Long, r As Long
    Dim guid As String
    Dim opened As Boolean
    Dim diaVal As Variant

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    Set wb = Workbooks.Open(Filename:=path, ReadOnly:=True, Local:=True)
    opened = True
    Set ws = PickBestSheet(wb, "bolt,report")
    FindHeaderRowAndGuidCol ws, headerRow, guidCol
    diaCol = FindColByNames(ws, headerRow, Array("BOLT_DIA", "Bolt_Dia", "Diameter", "DIA"))
    stdCol = FindColByNames(ws, headerRow, Array("BOLT_STD", "Bolt_Std", "Standard"))
    lastRow = ws.Cells(ws.Rows.Count, guidCol).End(xlUp).Row

    For r = headerRow + 1 To lastRow
        ' Trim spaces on bolt GUIDs
        guid = NormalizeGuid(ws.Cells(r, guidCol).Value)
        If InStr(guid, "-") > 0 Then
            diaVal = ""
            If diaCol > 0 Then
                diaVal = ws.Cells(r, diaCol).Value
                If VarType(diaVal) = vbString Then
                    diaVal = Trim$(CStr(diaVal))
                    If IsNumeric(diaVal) Then diaVal = CLng(diaVal)
                End If
            End If
            Dim std As String
            std = ""
            If stdCol > 0 Then std = Trim$(CStr(ws.Cells(r, stdCol).Value & ""))
            If Not dict.Exists(guid) Then
                dict.Add guid, Array(diaVal, std)
            End If
        End If
    Next r

    If opened Then wb.Close SaveChanges:=False
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
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "Guid:\s*([0-9a-fA-F-]{36})\s*<->\s*Guid:\s*([0-9a-fA-F-]{36})"

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

Private Function ClassifyAndWrite(ByVal clashes As Collection, ByVal parts As Object, ByVal bolts As Object) As Object
    Dim counts As Object
    Dim i As Long
    Dim item As Variant
    Dim g1 As String, g2 As String
    Dim k1 As String, k2 As String
    Dim isInside As Boolean
    Dim wsPP As Worksheet, wsPB As Worksheet, wsBB As Worksheet, wsDup As Worksheet, wsU As Worksheet
    Dim rPP As Long, rPB As Long, rBB As Long, rDup As Long, rU As Long

    Set counts = CreateObject("Scripting.Dictionary")
    counts.Add SHEET_PART_PART, 0
    counts.Add SHEET_PART_BOLT, 0
    counts.Add SHEET_BOLT_BOLT, 0
    counts.Add SHEET_DUP, 0
    counts.Add SHEET_UNRESOLVED, 0

    Set wsPP = ResetOutputSheet(SHEET_PART_PART, True)
    Set wsPB = ResetOutputSheet(SHEET_PART_BOLT, True)
    Set wsBB = ResetOutputSheet(SHEET_BOLT_BOLT, True)
    Set wsDup = ResetOutputSheet(SHEET_DUP, True)
    Set wsU = ResetOutputSheet(SHEET_UNRESOLVED, False)
    wsU.Range("A1").Value = "GUID1"
    wsU.Range("B1").Value = "GUID2"
    wsU.Range("C1").Value = "Kind1"
    wsU.Range("D1").Value = "Kind2"
    wsU.Range("E1").Value = "Log line"
    wsU.Range("A1:E1").Font.Bold = True

    rPP = 2: rPB = 2: rBB = 2: rDup = 2: rU = 2

    For i = 1 To clashes.Count
        item = clashes(i)
        g1 = item(0)
        g2 = item(1)
        isInside = item(2)
        k1 = KindOf(g1, parts, bolts)
        k2 = KindOf(g2, parts, bolts)

        If isInside Then
            WriteRow wsDup, rDup, FormatGuid(g1), FormatGuid(g2), AttrFor(g1, parts, bolts), AttrFor(g2, parts, bolts)
            rDup = rDup + 1
            counts(SHEET_DUP) = counts(SHEET_DUP) + 1
        Else
            ' Part-bolt: part first
            If k1 = "bolt" And k2 = "part" Then
                Dim tg As String, tk As String
                tg = g1: g1 = g2: g2 = tg
                tk = k1: k1 = k2: k2 = tk
            End If

            If k1 = "part" And k2 = "part" Then
                WriteRow wsPP, rPP, FormatGuid(g1), FormatGuid(g2), AttrFor(g1, parts, bolts), AttrFor(g2, parts, bolts)
                rPP = rPP + 1
                counts(SHEET_PART_PART) = counts(SHEET_PART_PART) + 1
            ElseIf k1 = "part" And k2 = "bolt" Then
                WriteRow wsPB, rPB, FormatGuid(g1), FormatGuid(g2), AttrFor(g1, parts, bolts), AttrFor(g2, parts, bolts)
                rPB = rPB + 1
                counts(SHEET_PART_BOLT) = counts(SHEET_PART_BOLT) + 1
            ElseIf k1 = "bolt" And k2 = "bolt" Then
                WriteRow wsBB, rBB, FormatGuid(g1), FormatGuid(g2), AttrFor(g1, parts, bolts), AttrFor(g2, parts, bolts)
                rBB = rBB + 1
                counts(SHEET_BOLT_BOLT) = counts(SHEET_BOLT_BOLT) + 1
            Else
                wsU.Cells(rU, 1).Value = FormatGuid(g1)
                wsU.Cells(rU, 2).Value = FormatGuid(g2)
                wsU.Cells(rU, 3).Value = k1
                wsU.Cells(rU, 4).Value = k2
                wsU.Cells(rU, 5).Value = item(3)
                rU = rU + 1
                counts(SHEET_UNRESOLVED) = counts(SHEET_UNRESOLVED) + 1
            End If
        End If
    Next i

    Set ClassifyAndWrite = counts
End Function

Private Sub WriteRow(ByVal ws As Worksheet, ByVal r As Long, ByVal g1 As String, ByVal g2 As String, ByVal a1 As Variant, ByVal a2 As Variant)
    ws.Cells(r, 1).Value = g1
    ws.Cells(r, 2).Value = g2
    ws.Cells(r, 3).Value = " "
    ws.Cells(r, 4).Value = a1
    ws.Cells(r, 5).Value = " "
    ws.Cells(r, 6).Value = a2
    ws.Cells(r, 7).Value = "" ' Notes (manual review) — matches sample final-tab style
End Sub

Private Function ResetOutputSheet(ByVal sheetName As String, ByVal withGuidHeader As Boolean) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = sheetName
    If withGuidHeader Then
        ws.Range("A1").Value = "GUID1"
        ws.Range("B1").Value = "GUID2"
        ws.Range("A1:B1").Font.Bold = True
    End If
    ws.Columns("A").ColumnWidth = 42
    ws.Columns("B").ColumnWidth = 42
    ws.Columns("C").ColumnWidth = 3
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 3
    ws.Columns("F").ColumnWidth = 18
    ws.Columns("G").ColumnWidth = 40
    Set ResetOutputSheet = ws
End Function

Private Sub WriteSummary(ByVal u As String, ByVal b As String, ByVal c As String, _
    ByVal partCount As Long, ByVal boltCount As Long, ByVal clashCount As Long, ByVal counts As Object)

    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(SHEET_SUMMARY).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(SHEET_CONTROL))
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
