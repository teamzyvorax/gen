Attribute VB_Name = "ZYVORA_Legacy"
Option Explicit

' ZYVORA Tools - Excel 2013 Legacy Add-in
' Import this module into a blank workbook and save it as ZYVORA_Tools.xlam.
' The installer script can automate this on a PC with Excel installed.

' ---------------- FORMULAS MASTER ----------------
Public Sub ZYVORA_InsertVLookup()
    ActiveCell.Formula = "=VLOOKUP(A1,B:C,2,FALSE)"
End Sub

Public Sub ZYVORA_InsertIfError()
    ActiveCell.Formula = "=IFERROR(A1,"")"
End Sub

Public Sub ZYVORA_InsertLeft()
    ActiveCell.Formula = "=LEFT(A1,2)"
End Sub

Public Sub ZYVORA_InsertRight()
    ActiveCell.Formula = "=RIGHT(A1,3)"
End Sub

Public Sub ZYVORA_InsertMid()
    ActiveCell.Formula = "=MID(A1,2,3)"
End Sub

Public Sub ZYVORA_InsertSubstitute()
    ActiveCell.Formula = "=SUBSTITUTE(A1,""."","""")"
End Sub

' ---------------- CODE GENERATOR ----------------
Public Sub ZYVORA_CodeFormula()
    Dim p As String, a As String, b As String, n As Long
    p = InputBox("Prefix:", "ZYVORA Code Generator", "LG")
    a = InputBox("First cell (example L21):", "ZYVORA Code Generator", "L21")
    b = InputBox("Second cell (example M21):", "ZYVORA Code Generator", "M21")
    n = Val(InputBox("Characters from second cell:", "ZYVORA Code Generator", "3"))
    If n < 1 Then n = 3
    ActiveCell.Formula = "=LEFT(""" & Replace(p, """", """"") & """," & Len(p) & ")&" & a & "&LEFT(SUBSTITUTE(" & b & ",""."","""")," & n & ")"
End Sub

' ---------------- TEXT TOOLS ----------------
Public Sub ZYVORA_TrimSelected(): ZYVORA_TextFormula "TRIM": End Sub
Public Sub ZYVORA_UpperSelected(): ZYVORA_TextFormula "UPPER": End Sub
Public Sub ZYVORA_LowerSelected(): ZYVORA_TextFormula "LOWER": End Sub
Public Sub ZYVORA_ProperSelected(): ZYVORA_TextFormula "PROPER": End Sub
Public Sub ZYVORA_CleanSelected(): ZYVORA_TextFormula "CLEAN": End Sub
Public Sub ZYVORA_LenSelected(): ZYVORA_TextFormula "LEN": End Sub

Private Sub ZYVORA_TextFormula(ByVal fn As String)
    Dim src As Range, out As Range, c As Range
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set src = Selection
    Set out = src.Offset(0, src.Columns.Count)
    For Each c In src.Cells
        out.Cells(c.Row - src.Row + 1, c.Column - src.Column + 1).Formula = "=" & fn & "(" & c.Address(False, False) & ")"
    Next c
End Sub

' ---------------- DATA TOOLS ----------------
Public Sub ZYVORA_AutoFit()
    If TypeName(Selection) = "Range" Then Selection.EntireColumn.AutoFit
End Sub

Public Sub ZYVORA_ClearFormats()
    If TypeName(Selection) = "Range" Then Selection.ClearFormats
End Sub

Public Sub ZYVORA_FillDown()
    If TypeName(Selection) = "Range" Then Selection.FillDown
End Sub

Public Sub ZYVORA_SortAZ()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Selection.Sort Key1:=Selection.Cells(1, 1), Order1:=xlAscending, Header:=xlGuess
End Sub

Public Sub ZYVORA_RemoveDuplicates()
    Dim i As Long, cols() As Variant
    If TypeName(Selection) <> "Range" Then Exit Sub
    ReDim cols(1 To Selection.Columns.Count)
    For i = 1 To Selection.Columns.Count: cols(i) = i: Next i
    Selection.RemoveDuplicates Columns:=cols, Header:=xlYes
End Sub

Public Sub ZYVORA_FreezeTopRow()
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True
End Sub

' ---------------- RAPNET TOOLS ----------------
Public Sub ZYVORA_RapNetHeader()
    Dim r As Range
    Set r = ActiveCell
    r.Value = "SKU"
    r.Offset(0, 1).Value = "Description"
    r.Offset(0, 2).Value = "Quantity"
    r.Offset(0, 3).Value = "Price"
    r.Offset(0, 4).Value = "Status"
    r.Offset(0, 5).Value = "Reference"
    r.Resize(1, 6).Font.Bold = True
End Sub

Public Sub ZYVORA_RapNetCleanSelected()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Selection.Replace What:=Chr(160), Replacement:=" ", LookAt:=xlPart
    Selection.Replace What:="  ", Replacement:=" ", LookAt:=xlPart
    Selection.Replace What:=",", Replacement:="", LookAt:=xlPart
    Selection.EntireColumn.AutoFit
End Sub

Public Sub ZYVORA_RapNetExportReady()
    Dim ws As Worksheet, out As Worksheet, src As Range
    Set ws = ActiveSheet
    Set src = Selection
    Set out = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    out.Name = "RapNet_Export_" & Format(Now, "hhmmss")
    src.Copy Destination:=out.Range("A1")
    out.UsedRange.EntireColumn.AutoFit
    MsgBox "RapNet-ready copy created: " & out.Name, vbInformation, "ZYVORA RapNet"
End Sub

' ---------------- BARCODE TOOLS ----------------
Public Sub ZYVORA_BarcodeCode128Text()
    Dim c As Range
    If TypeName(Selection) <> "Range" Then Exit Sub
    For Each c In Selection.Cells
        If Len(CStr(c.Value)) > 0 Then
            c.Offset(0, 1).Value = "*" & CStr(c.Value) & "*"
        End If
    Next c
    MsgBox "Barcode text prepared. For a scannable barcode, install a Code 39/Code 128 barcode font and apply it to the generated cells.", vbInformation, "ZYVORA Barcode"
End Sub

Public Sub ZYVORA_BarcodeSheet()
    Dim ws As Worksheet
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.Name = "Barcode_" & Format(Now, "hhmmss")
    ws.Range("A1").Value = "ZYVORA Barcode Sheet"
    ws.Range("A1").Font.Bold = True
    ws.Range("A3").Value = "Enter SKU / barcode values in column A"
    ws.Range("B3").Value = "Code 39 text"
    ws.Columns("A:B").ColumnWidth = 24
    MsgBox "Barcode sheet created. Put values in column A, then run Barcode Code 39 Text.", vbInformation, "ZYVORA Barcode"
End Sub

' ---------------- AI TOOLS ----------------
Public Sub ZYVORA_AIPrompt()
    Dim q As String
    q = InputBox("What do you want the AI prompt to do?", "ZYVORA AI Tools", "Clean and standardize this Excel data")
    If Len(q) = 0 Then Exit Sub
    ActiveCell.Value = "Prompt: " & q & vbCrLf & "Use the selected Excel range as context. Return a concise, Excel-ready result."
End Sub

' ---------------- CUSTOM FORMULA BUILDER ----------------
Public Sub ZYVORA_CustomFormulaBuilder()
    Dim fn As String, ref As String, f As String
    fn = InputBox("Function name (example LEFT, TRIM, UPPER):", "ZYVORA Formula Builder", "TRIM")
    ref = InputBox("Cell reference:", "ZYVORA Formula Builder", "A1")
    If Len(fn) = 0 Or Len(ref) = 0 Then Exit Sub
    f = "=" & fn & "(" & ref & ")"
    ActiveCell.Formula = f
End Sub

' ---------------- SETTINGS & HELP ----------------
Public Sub ZYVORA_Help()
    MsgBox "ZYVORA Tools - Excel 2013 Legacy\n\nUse the ZYVORA buttons/macros for formulas, text cleanup, data tools, RapNet formatting, barcode preparation, AI prompt helpers and custom formulas.", vbInformation, "ZYVORA Tools"
End Sub

Public Sub ZYVORA_ShowVersion()
    MsgBox "ZYVORA Tools | Excel 2013 Legacy | RapNet + Barcode", vbInformation, "ZYVORA Tools"
End Sub
