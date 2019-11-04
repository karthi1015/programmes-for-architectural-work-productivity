''' RAW_DATA '''
Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo exitsub

    If Target.Rows > 0 Then
    Application.EnableEvents = False

    Call raw_data

    End If
    Application.EnableEvents = True

exitsub:
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo exitsub

    If Target.Rows > 0 Then
    Application.EnableEvents = False

    Call raw_data

    End If

    Application.EnableEvents = True

exitsub:
    Application.EnableEvents = True
End Sub


''' MODULE '''
Sub Refresh()

    Call Protect

    Worksheets("RAW_DATA").Range("B34:KA300").ClearFormats

    Dim rng_1 As Range
    Dim rng_2 As Range
    Dim rng_3 As Range
    Dim rng_4 As Range
    Dim rng_5 As Range
    Dim rng_6 As Range
    Dim rng_7 As Range
    Dim rng_8 As Range
    Dim rng_9 As Range
    Dim rng_10 As Range

    Set rng_1 = ThisWorkbook.Sheets("RAW_DATA").Range("E7:KA9")
    Set rng_2 = ThisWorkbook.Sheets("RAW_DATA").Range("E12:KA21")
    Set rng_3 = ThisWorkbook.Sheets("RAW_DATA").Range("E22:KA22")
    Set rng_4 = ThisWorkbook.Sheets("RAW_DATA").Range("B12:D21")
    Set rng_5 = ThisWorkbook.Sheets("RAW_DATA").Range("E25:KA26")
    Set rng_6 = ThisWorkbook.Sheets("RAW_DATA").Range("E29:KA31")
    Set rng_7 = ThisWorkbook.Sheets("RAW_DATA").Range("B34:B300")
    Set rng_8 = ThisWorkbook.Sheets("RAW_DATA").Range("C34:C300")
    Set rng_9 = ThisWorkbook.Sheets("RAW_DATA").Range("D34:D300")
    Set rng_10 = ThisWorkbook.Sheets("RAW_DATA").Range("E34:KA300")
    Set rng_11 = ThisWorkbook.Sheets("PRINT").Range("A1:AJ1300")

    rng_1.Borders(xlInsideVertical).LineStyle = xlNone
    rng_1.Borders(xlInsideHorizontal).LineStyle = xlNone
    rng_2.Borders(xlInsideVertical).LineStyle = xlNone
    rng_2.Borders(xlInsideHorizontal).LineStyle = xlNone
    rng_3.Borders(xlInsideVertical).LineStyle = xlNone
    rng_3.Borders(xlInsideHorizontal).LineStyle = xlNone
    rng_4.Borders(xlInsideVertical).LineStyle = xlNone
    rng_4.Borders(xlInsideHorizontal).LineStyle = xlNone
    rng_5.Borders(xlInsideVertical).LineStyle = xlNone
    rng_5.Borders(xlInsideHorizontal).LineStyle = xlNone
    rng_6.Borders(xlInsideVertical).LineStyle = xlNone
    rng_6.Borders(xlInsideHorizontal).LineStyle = xlNone
    rng_7.Borders(xlInsideVertical).LineStyle = xlNone
    rng_7.Borders(xlInsideHorizontal).LineStyle = xlNone
    rng_8.Borders(xlInsideVertical).LineStyle = xlNone
    rng_8.Borders(xlInsideHorizontal).LineStyle = xlNone
    rng_9.Borders(xlInsideVertical).LineStyle = xlNone
    rng_9.Borders(xlInsideHorizontal).LineStyle = xlNone
    rng_10.Borders(xlInsideVertical).LineStyle = xlNone
    rng_10.Borders(xlInsideHorizontal).LineStyle = xlNone

    With rng_1.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng_1.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng_1.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With ThisWorkbook.Sheets("RAW_DATA").Rows(12).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With ThisWorkbook.Sheets("RAW_DATA").Rows(22).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With ThisWorkbook.Sheets("RAW_DATA").Rows(22).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With ThisWorkbook.Sheets("RAW_DATA").Range("E12:E22").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng_5.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng_5.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng_5.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng_6.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With rng_6.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng_6.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng_6.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng_7.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng_7.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng_7.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With rng_8.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng_8.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With rng_9.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng_9.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng_9.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng_9.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With rng_10.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
    End With
    With rng_10.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    With rng_1.Font
        .Name = "Avenir LT Std 45 Book"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With rng_2.Font
        .Name = "Avenir LT Std 45 Book"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With rng_3.Font
        .Name = "Avenir LT Std 45 Book"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With rng_4.Font
        .Name = "Avenir LT Std 45 Book"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With rng_5.Font
        .Name = "Avenir LT Std 45 Book"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With rng_6.Font
        .Name = "Avenir LT Std 45 Book"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With rng_7.Font
        .Name = "Avenir LT Std 45 Book"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With rng_7
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Locked = False
        .FormulaHidden = False
    End With
    With rng_8.Font
        .Name = "Avenir LT Std 45 Book"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With rng_8
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Locked = False
        .FormulaHidden = False
    End With
    With rng_9.Font
        .Name = "Avenir LT Std 45 Book"
        .Size = 14
        .Bold = True
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With rng_9
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Locked = False
        .FormulaHidden = False
    End With
    With rng_10.Font
        .Name = "Avenir LT Std 45 Book"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With rng_10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Locked = False
        .FormulaHidden = False
    End With

    ThisWorkbook.Sheets("TERMINAL").Range("B34:B300").Value = ThisWorkbook.Sheets("RAW_DATA").Range("B34:B300").Value
    ThisWorkbook.Sheets("TERMINAL").Range("C34:B300").Value = ThisWorkbook.Sheets("RAW_DATA").Range("C34:B300").Value
    ThisWorkbook.Sheets("TERMINAL").Range("E34:KA300").Value = ThisWorkbook.Sheets("RAW_DATA").Range("E34:KA300").Value
    ThisWorkbook.Sheets("TERMINAL").Range("D34:D300").Cells.Replace "#N/A", " ", xlWhole
    ThisWorkbook.Sheets("RAW_DATA").Range("D34:D300").Value = ThisWorkbook.Sheets("TERMINAL").Range("D34:D300").Value
    ThisWorkbook.Sheets("TERMINAL").Range("E7:KA9").Value = ThisWorkbook.Sheets("RAW_DATA").Range("E7:KA9").Value
    ThisWorkbook.Sheets("TERMINAL").Range("E12:KA21").Value = ThisWorkbook.Sheets("RAW_DATA").Range("E12:KA21").Value
    ThisWorkbook.Sheets("TERMINAL").Range("E22:KA22").Value = ThisWorkbook.Sheets("RAW_DATA").Range("E22:KA22").Value


    ThisWorkbook.Sheets("PRINT_TEMP").Cells(1, 20).Value = ThisWorkbook.Sheets("RAW_DATA").Cells(2, 3).Value
    ThisWorkbook.Sheets("PRINT_TEMP").Cells(3, 20).Value = ThisWorkbook.Sheets("RAW_DATA").Cells(4, 3).Value
    ThisWorkbook.Sheets("PRINT_TEMP").Range("S6:AJ8").Value = ThisWorkbook.Sheets("RAW_DATA").Range("E7:KA9").Value
    ThisWorkbook.Sheets("PRINT_TEMP").Range("A48:A57").Value = ThisWorkbook.Sheets("RAW_DATA").Range("B12:B21").Value
    ThisWorkbook.Sheets("PRINT_TEMP").Range("S48:AJ57").Value = ThisWorkbook.Sheets("RAW_DATA").Range("E12:KA21").Value
    ThisWorkbook.Sheets("PRINT_TEMP").Range("S59:AJ59").Value = ThisWorkbook.Sheets("RAW_DATA").Range("E22:KA22").Value
    ThisWorkbook.Sheets("PRINT_TEMP").Range("S62:AJ63").Value = ThisWorkbook.Sheets("RAW_DATA").Range("E25:KA26").Value
    ThisWorkbook.Sheets("PRINT_TEMP").Range("S66:AJ67").Value = ThisWorkbook.Sheets("RAW_DATA").Range("E29:KA30").Value

    For x = 23 To 300

    If ThisWorkbook.Sheets("TERMINAL").Cells(1, x).Value = 0 Then Exit For
        ThisWorkbook.Sheets("PRINT_TEMP").Range("S6:AJ8").Value = ThisWorkbook.Sheets("RAW_DATA").Range(Sheets("RAW_DATA").Cells(7, x - 17), Sheets("RAW_DATA").Cells(9, x)).Value
        ThisWorkbook.Sheets("PRINT_TEMP").Range("S48:AJ57").Value = ThisWorkbook.Sheets("RAW_DATA").Range(Sheets("RAW_DATA").Cells(12, x - 17), Sheets("RAW_DATA").Cells(21, x)).Value
        ThisWorkbook.Sheets("PRINT_TEMP").Range("S59:AJ59").Value = ThisWorkbook.Sheets("RAW_DATA").Range(Sheets("RAW_DATA").Cells(22, x - 17), Sheets("RAW_DATA").Cells(22, x)).Value
        ThisWorkbook.Sheets("PRINT_TEMP").Range("S62:AJ63").Value = ThisWorkbook.Sheets("RAW_DATA").Range(Sheets("RAW_DATA").Cells(25, x - 17), Sheets("RAW_DATA").Cells(26, x)).Value
        ThisWorkbook.Sheets("PRINT_TEMP").Range("S66:AJ67").Value = ThisWorkbook.Sheets("RAW_DATA").Range(Sheets("RAW_DATA").Cells(29, x - 17), Sheets("RAW_DATA").Cells(30, x)).Value
    Next x

    ThisWorkbook.Sheets("PRINT").Cells.Clear


    Call pagenumber_set
    Call page_set

    With rng_11
        .VerticalAlignment = xlTop
    End With

    Call pagebreak
    Call project_align

End Sub

Sub PrintOut()

    On Error GoTo exitsub
    Call Refresh
    
    wb_name = Replace(ThisWorkbook.Name, ".xlsm", "")
    
'    ThisWorkbook.Sheets("PRINT").PrintOut copies:=1, ActivePrinter:="PDFCreator", printtofile:=True, collate:=True, prtofilename:=PSFileName

'    ThisWorkbook.Sheets("PRINT").ExportAsFixedFormat Type:=xlTypePDF, Filename:=ThisWorkbook.Path & "\" & "Drawing Register Engenium", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    
    ThisWorkbook.Sheets("PRINT").ExportAsFixedFormat Type:=xlTypePDF, Filename:=ThisWorkbook.Path & "\" & wb_name, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    MsgBox "PDF has been sucessfully saved"
    Exit Sub

'    ThisWorkbook.Sheets("PRINT").Select
'    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ThisWorkbook.Path & "\" & "Drawing Register Engenium", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
'    ThisWorkbook.Worksheets("RAW_DATA").Activate
'
exitsub:
    MsgBox "PDF cannot be created. The file may be already in use."
    Application.EnableEvents = True
    
End Sub

Sub page_set()

    Dim ax As Integer
    Dim ay As Integer
    Dim bx As Integer
    Dim by As Integer

    For i = 1 To 18
        ax = 11 + 67 * (i - 1)
        ay = 45 + 67 * (i - 1)
        bx = (34 * i) + (i - 1)
        by = bx + 34

        ThisWorkbook.Sheets("PRINT").Range("A" & ax & ":" & "A" & ay).Value = ThisWorkbook.Sheets("RAW_DATA").Range("B" & bx & ":" & "B" & by).Value
        ThisWorkbook.Sheets("PRINT").Range("E" & ax & ":" & "E" & ay).Value = ThisWorkbook.Sheets("RAW_DATA").Range("C" & bx & ":" & "C" & by).Value
        ThisWorkbook.Sheets("PRINT").Range("R" & ax & ":" & "R" & ay).Value = ThisWorkbook.Sheets("RAW_DATA").Range("D" & bx & ":" & "D" & by).Value
        ThisWorkbook.Sheets("PRINT").Range("S" & ax & ":" & "AJ" & ay).Value = ThisWorkbook.Sheets("RAW_DATA").Range("E" & bx & ":" & "V" & by).Value

        For E = 23 To 300

        If ThisWorkbook.Sheets("TERMINAL").Cells(1, E).Value = 0 Then Exit For

            ThisWorkbook.Sheets("PRINT").Range("S" & ax & ":" & "AJ" & ay).Value = ThisWorkbook.Sheets("RAW_DATA").Range(Sheets("RAW_DATA").Cells(bx, E - 17), Sheets("RAW_DATA").Cells(by, E)).Value

        Next E

    Next i

End Sub

Sub project_align()

    Dim pa As Integer
    
    For pa_num = 0 To 4
        num = (pa_num * 67) + 1
        With ThisWorkbook.Sheets("PRINT")
            .Cells(num, 19).VerticalAlignment = xlCenter
            .Cells(num, 20).VerticalAlignment = xlCenter
            .Cells(num, 19).Offset(1, 0).VerticalAlignment = xlCenter
            .Cells(num, 20).Offset(1, 0).VerticalAlignment = xlCenter
        End With
    Next pa_num

End Sub

Sub pagenumber_set()

    Dim tx As Integer
    Dim ty As Integer
    Dim trng As Range
    Dim cell As Range
    Dim pt As Integer

    For p = 1 To 4
        tx = (34 * p) + (p - 1)

        If ThisWorkbook.Sheets("TERMINAL").Range("A" & tx).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 1).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 2).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 3).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 4).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 5).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 6).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 7).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 8).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 9).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 10).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 11).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 12).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 13).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 14).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 15).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 16).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 17).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 18).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 19).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 20).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 21).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 22).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 23).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 24).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 25).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 26).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 27).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 28).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 29).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 30).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 31).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 32).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 33).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        ElseIf ThisWorkbook.Sheets("TERMINAL").Range("A" & tx + 34).Value > 0 Then
            pt = 1 + 67 * (p - 1)
            Application.CutCopyMode = False
            ThisWorkbook.Sheets("PRINT_TEMP").Range("A1:AJ67").Copy
            ThisWorkbook.Sheets("PRINT").Range("A" & pt).PasteSpecial xlPasteAll
        End If
    Next p
End Sub

Sub raw_data()

    Call Protect

        Dim rng_1 As Range
        Dim rng_2 As Range
        Dim rng_3 As Range
        Dim rng_4 As Range
        Dim rng_5 As Range
        Dim rng_6 As Range
        Dim rng_7 As Range
        Dim rng_8 As Range
        Dim rng_9 As Range
        Dim rng_10 As Range

        Set rng_1 = ThisWorkbook.Sheets("RAW_DATA").Range("E7:KA9")
        Set rng_2 = ThisWorkbook.Sheets("RAW_DATA").Range("E12:KA21")
        Set rng_3 = ThisWorkbook.Sheets("RAW_DATA").Range("E22:KA22")
        Set rng_4 = ThisWorkbook.Sheets("RAW_DATA").Range("B12:D21")
        Set rng_5 = ThisWorkbook.Sheets("RAW_DATA").Range("E25:KA26")
        Set rng_6 = ThisWorkbook.Sheets("RAW_DATA").Range("E29:KA31")
        Set rng_7 = ThisWorkbook.Sheets("RAW_DATA").Range("B34:B300")
        Set rng_8 = ThisWorkbook.Sheets("RAW_DATA").Range("C34:C300")
        Set rng_9 = ThisWorkbook.Sheets("RAW_DATA").Range("D34:D300")
        Set rng_10 = ThisWorkbook.Sheets("RAW_DATA").Range("E34:KA300")

        rng_1.Borders(xlInsideVertical).LineStyle = xlNone
        rng_1.Borders(xlInsideHorizontal).LineStyle = xlNone
        rng_2.Borders(xlInsideVertical).LineStyle = xlNone
        rng_2.Borders(xlInsideHorizontal).LineStyle = xlNone
        rng_3.Borders(xlInsideVertical).LineStyle = xlNone
        rng_3.Borders(xlInsideHorizontal).LineStyle = xlNone
        rng_4.Borders(xlInsideVertical).LineStyle = xlNone
        rng_4.Borders(xlInsideHorizontal).LineStyle = xlNone
        rng_5.Borders(xlInsideVertical).LineStyle = xlNone
        rng_5.Borders(xlInsideHorizontal).LineStyle = xlNone
        rng_6.Borders(xlInsideVertical).LineStyle = xlNone
        rng_6.Borders(xlInsideHorizontal).LineStyle = xlNone
        rng_7.Borders(xlInsideVertical).LineStyle = xlNone
        rng_7.Borders(xlInsideHorizontal).LineStyle = xlNone
        rng_8.Borders(xlInsideVertical).LineStyle = xlNone
        rng_8.Borders(xlInsideHorizontal).LineStyle = xlNone
        rng_9.Borders(xlInsideVertical).LineStyle = xlNone
        rng_9.Borders(xlInsideHorizontal).LineStyle = xlNone
        rng_10.Borders(xlInsideVertical).LineStyle = xlNone
        rng_10.Borders(xlInsideHorizontal).LineStyle = xlNone

        With rng_7.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With rng_7.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With rng_8.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With rng_9.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With rng_9.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With rng_9.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With rng_10.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With

        ThisWorkbook.Sheets("TERMINAL").Range("B34:B300").Value = ThisWorkbook.Sheets("RAW_DATA").Range("B34:B300").Value
        ThisWorkbook.Sheets("TERMINAL").Range("C34:B300").Value = ThisWorkbook.Sheets("RAW_DATA").Range("C34:B300").Value
        ThisWorkbook.Sheets("TERMINAL").Range("E34:KA300").Value = ThisWorkbook.Sheets("RAW_DATA").Range("E34:KA300").Value
        ThisWorkbook.Sheets("TERMINAL").Range("D34:D300").Cells.Replace "#N/A", " ", xlWhole
        ThisWorkbook.Sheets("RAW_DATA").Range("D34:D300").Value = ThisWorkbook.Sheets("TERMINAL").Range("D34:D300").Value
        ThisWorkbook.Sheets("TERMINAL").Range("E7:KA9").Value = ThisWorkbook.Sheets("RAW_DATA").Range("E7:KA9").Value
        ThisWorkbook.Sheets("TERMINAL").Range("E12:KA21").Value = ThisWorkbook.Sheets("RAW_DATA").Range("E12:KA21").Value
        ThisWorkbook.Sheets("TERMINAL").Range("E22:KA22").Value = ThisWorkbook.Sheets("RAW_DATA").Range("E22:KA22").Value
End Sub

Sub Protect()

    Worksheets("RAW_DATA").Protect "Engenium", DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowInsertingColumns:=True, AllowInsertingRows:=True, _
        AllowDeletingColumns:=True, AllowDeletingRows:=True, userinterfaceonly:=True
    Worksheets("RAW_DATA").EnableSelection = xlUnlockedCells

    Worksheets("PRINT").Protect "Engenium", userinterfaceonly:=True
    Worksheets("PRINT_TEMP").Protect "Engenium", userinterfaceonly:=True
    Worksheets("TERMINAL").Protect "Engenium", userinterfaceonly:=True

End Sub

Sub UnProtect()

    For i = 1 To Sheets.Count
        Sheets(i).UnProtect "Engenium"
    Next i

End Sub

Sub pagebreak()

    Dim p_interval As Integer
    Dim p_chec As Integer

    ThisWorkbook.Sheets("PRINT").ResetAllPageBreaks
    ThisWorkbook.Sheets("PRINT").VPageBreaks.Add before:=Columns("AK")

    For i = 0 To 4
        p_chec = 11 + (67 * i)
        If ThisWorkbook.Sheets("PRINT").Cells(p_chec, 1) > 0 Then
            p_interval = p_chec + 57
            ThisWorkbook.Sheets("PRINT").HPageBreaks.Add before:=Rows(p_interval)
        End If
    Next

    ThisWorkbook.Sheets("PRINT").PageSetup.PrintArea = ""

End Sub

Sub structure()
    On Error GoTo exitsub
    
    Dim rng As Range
    Set rng = Application.InputBox("Select a range", "Obtain Range Object", Type:=8)

    ThisWorkbook.Sheets("Structural_Drawinglist").Range("A4:B300").Value = rng.Value
    ThisWorkbook.Sheets("Structural_Drawinglist").Range("A4:B300").Cells.Replace "#N/A", " ", xlWhole

exitsub:
    Application.EnableEvents = True
End Sub

Sub panel()
    On Error GoTo exitsub
    
    Dim rng As Range
    Set rng = Application.InputBox("Select a range", "Obtain Range Object", Type:=8)

    ThisWorkbook.Sheets("Panels_Drawinglist").Range("A4:B300").Value = rng.Value
    ThisWorkbook.Sheets("Panels_Drawinglist").Range("A4:B300").Cells.Replace "#N/A", " ", xlWhole

exitsub:
    Application.EnableEvents = True
End Sub

