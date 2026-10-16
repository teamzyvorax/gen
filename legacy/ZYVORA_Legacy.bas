Attribute VB_Name = "ZYVORA_Legacy"
Option Explicit

' ZYVORA Legacy Compatibility Pack
' Intended for Excel 2007/2010 and other builds where Office.js web add-ins are unavailable.
' Import this module into a workbook, then save the workbook as an Excel Add-In (*.xlam).

Public Sub ZYVORA_CodeFormula()
    Dim p As String, a As String, b As String, n As Long
    p = InputBox("Prefix:", "ZYVORA Code Generator", "LG")
    a = InputBox("First cell (example L21):", "ZYVORA Code Generator", "L21")
    b = InputBox("Second cell (example M21):", "ZYVORA Code Generator", "M21")
    n = Val(InputBox("Characters from second cell:", "ZYVORA Code Generator", "3"))
    If n < 1 Then n = 3
    ActiveCell.Formula = "=LEFT(""" & Replace(p, """", """"") & """," & Len(p) & ")&" & a & "&LEFT(SUBSTITUTE(" & b & ",""."","""")," & n & ")"
End Sub

Public Sub ZYVORA_TrimSelected()
    ZYVORA_TextFormula "TRIM"
End Sub

Public Sub ZYVORA_UpperSelected()
    ZYVORA_TextFormula "UPPER"
End Sub

Public Sub ZYVORA_LowerSelected()
    ZYVORA_TextFormula "LOWER"
End Sub

Public Sub ZYVORA_ProperSelected()
    ZYVORA_TextFormula "PROPER"
End Sub

Public Sub ZYVORA_CleanSelected()
    ZYVORA_TextFormula "CLEAN"
End Sub

Public Sub ZYVORA_LenSelected()
    ZYVORA_TextFormula "LEN"
End Sub

Private Sub ZYVORA_TextFormula(ByVal fn As String)
    Dim src As Range, out As Range, c As Range
    Set src = Selection
    Set out = src.Offset(0, src.Columns.Count)
    For Each c In src.Cells
        out.Cells(c.Row - src.Row + 1, c.Column - src.Column + 1).Formula = "=" & fn & "(" & c.Address(False, False) & ")"
    Next c
End Sub

Public Sub ZYVORA_AutoFit()
    Selection.EntireColumn.AutoFit
End Sub

Public Sub ZYVORA_ClearFormats()
    Selection.ClearFormats
End Sub

Public Sub ZYVORA_FillDown()
    Selection.FillDown
End Sub

Public Sub ZYVORA_SortAZ()
    If Selection.Columns.Count = 0 Then Exit Sub
    Selection.Sort Key1:=Selection.Cells(1, 1), Order1:=xlAscending, Header:=xlGuess
End Sub

Public Sub ZYVORA_RemoveDuplicates()
    Dim i As Long, cols() As Variant
    ReDim cols(1 To Selection.Columns.Count)
    For i = 1 To Selection.Columns.Count
        cols(i) = i
    Next i
    Selection.RemoveDuplicates Columns:=cols, Header:=xlYes
End Sub

Public Sub ZYVORA_FreezeTopRow()
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True
End Sub

Public Sub ZYVORA_InsertVLookup()
    ActiveCell.Formula = "=VLOOKUP(A1,B:C,2,FALSE)"
End Sub

Public Sub ZYVORA_InsertIfError()
    ActiveCell.Formula = "=IFERROR(A1,"")"
End Sub

Public Sub ZYVORA_InsertTextJoinLegacy()
    ' TEXTJOIN is not available in very old Excel. This uses simple concatenation.
    ActiveCell.Formula = "=A1&""-""&B1"
End Sub
