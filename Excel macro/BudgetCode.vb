' PROJECT_BUDGET '

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo exitsub

    If Target.Rows > 0 Then
    Application.EnableEvents = False
    
    Call Protect
    Call Project_Budget_M
    
    End If
    Application.EnableEvents = True

exitsub:
    Application.EnableEvents = True
    
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo exitsub

    If Target.Rows > 0 Then
    Application.EnableEvents = False

    Call Protect
    Call Project_Budget_M
    
    End If
    Application.EnableEvents = True

exitsub:
    Application.EnableEvents = True

End Sub


' SCOPE_CHANGE '

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo exitsub

    If Target.Rows > 0 Then
    Application.EnableEvents = False

    Call Scope_Change_M
    
    End If
    Application.EnableEvents = True

exitsub:
    Application.EnableEvents = True
    
End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo exitsub

    If Target.Rows > 0 Then
    Application.EnableEvents = False

    Call Scope_Change_M

    End If
    Application.EnableEvents = True

exitsub:
    Application.EnableEvents = True

End Sub

' REWORK '

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo exitsub

    If Target.Rows > 0 Then
    Application.EnableEvents = False

    Call Rework_M
    
    End If
    Application.EnableEvents = True

exitsub:
    Application.EnableEvents = True
    
End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo exitsub

    If Target.Rows > 0 Then
    Application.EnableEvents = False

    Call Rework_M

    End If
    Application.EnableEvents = True

exitsub:
    Application.EnableEvents = True

End Sub


' Module '

Sub PrintOut()
    On Error GoTo exitsub
    
    Call Project_Budget_M
    Call Scope_Change_M
    Call Rework_M

    Dim re_ver As Integer
    
    With ThisWorkbook.Sheets("Project_Budget").Cells.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With

    For y = 7 To 30
        If ThisWorkbook.Sheets("Project_Budget").Cells(y, 7).Value < 0 Then
            re_ver = 0
            For a = 8 To 14 Step 2
                re_ver = ThisWorkbook.Sheets("Project_Budget").Cells(y, a).Value + re_ver
                If ThisWorkbook.Sheets("Project_Budget").Cells(y, 2).Value - re_ver < 0 Then
                    With Cells(y, a).Font
                        .Color = -16776961
                        .TintAndShade = 0
                    End With
                End If
            Next a
            For b = 9 To 15 Step 2
                re_ver = ThisWorkbook.Sheets("Project_Budget").Cells(y, b).Value + re_ver
                If ThisWorkbook.Sheets("Project_Budget").Cells(y, 3).Value - re_ver < 0 Then
                    With Cells(y, b).Font
                        .Color = -16776961
                        .TintAndShade = 0
                    End With
                End If
            Next b
            For c = 17 To 29 Step 2
                re_ver = ThisWorkbook.Sheets("Project_Budget").Cells(y, c).Value + re_ver
                If ThisWorkbook.Sheets("Project_Budget").Cells(y, 2).Value - re_ver < 0 Then
                    With Cells(y, c).Font
                        .Color = -16776961
                        .TintAndShade = 0
                    End With
                End If
            Next c
            For d = 18 To 30 Step 2
                re_ver = ThisWorkbook.Sheets("Project_Budget").Cells(y, d).Value + re_ver
                If ThisWorkbook.Sheets("Project_Budget").Cells(y, 3).Value - re_ver < 0 Then
                    With Cells(y, d).Font
                        .Color = -16776961
                        .TintAndShade = 0
                    End With
                End If
            Next d
        End If
    Next y

    ThisWorkbook.Sheets("PRINT").Cells.Clear
    ThisWorkbook.Sheets("6T_PRINT_1").Cells.Clear
    ThisWorkbook.Sheets("6T_PRINT_2").Cells.Clear
    ThisWorkbook.Sheets("7T_PRINT_1").Cells.Clear
    ThisWorkbook.Sheets("7T_PRINT_2").Cells.Clear

    Dim brng_1 As Range
    Dim brng_2 As Range

    Dim bprng_1 As Range
    Dim bprng_2 As Range

    Dim srng_1_1 As Range
    Dim srng_2_1 As Range
    Dim srng_3_1 As Range
    Dim srng_4_1 As Range
    Dim srng_5_1 As Range
    Dim srng_6_1 As Range

    Dim prng_1_1 As Range
    Dim prng_2_1 As Range
    Dim prng_3_1 As Range
    Dim prng_4_1 As Range
    Dim prng_5_1 As Range
    Dim prng_6_1 As Range

    Dim srng_1_2 As Range
    Dim srng_2_2 As Range
    Dim srng_3_2 As Range
    Dim srng_4_2 As Range
    Dim srng_5_2 As Range
    Dim srng_6_2 As Range

    Dim prng_1_2 As Range
    Dim prng_2_2 As Range
    Dim prng_3_2 As Range
    Dim prng_4_2 As Range
    Dim prng_5_2 As Range
    Dim prng_6_2 As Range

    Dim rerng_1_1 As Range
    Dim rerng_2_1 As Range
    Dim rerng_3_1 As Range
    Dim rerng_4_1 As Range
    Dim rerng_5_1 As Range
    Dim rerng_6_1 As Range

    Dim reprng_1_1 As Range
    Dim reprng_2_1 As Range
    Dim reprng_3_1 As Range
    Dim reprng_4_1 As Range
    Dim reprng_5_1 As Range
    Dim reprng_6_1 As Range

    Dim rerng_1_2 As Range
    Dim rerng_2_2 As Range
    Dim rerng_3_2 As Range
    Dim rerng_4_2 As Range
    Dim rerng_5_2 As Range
    Dim rerng_6_2 As Range

    Dim reprng_1_2 As Range
    Dim reprng_2_2 As Range
    Dim reprng_3_2 As Range
    Dim reprng_4_2 As Range
    Dim reprng_5_2 As Range
    Dim reprng_6_2 As Range

    Set brng_1 = ThisWorkbook.Sheets("Project_Budget").Range("A1:O46")
    Set brng_2 = ThisWorkbook.Sheets("Project_Budget").Range("P1:AE46")

    Set bprng_1 = ThisWorkbook.Sheets("PRINT").Range("A1:O46")
    Set bprng_2 = ThisWorkbook.Sheets("PRINT").Range("P1:AE46")

    Set srng_1_1 = ThisWorkbook.Sheets("Scope_Change").Range("A1:L47")
    Set srng_2_1 = ThisWorkbook.Sheets("Scope_Change").Range("M1:AA47")
    Set srng_3_1 = ThisWorkbook.Sheets("Scope_Change").Range("AB1:AP47")
    Set srng_4_1 = ThisWorkbook.Sheets("Scope_Change").Range("AQ1:BE47")
    Set srng_5_1 = ThisWorkbook.Sheets("Scope_Change").Range("BF1:BT47")
    Set srng_6_1 = ThisWorkbook.Sheets("Scope_Change").Range("BU1:CI47")

    Set prng_1_1 = ThisWorkbook.Sheets("6T_PRINT_1").Range("A1:L47")
    Set prng_2_1 = ThisWorkbook.Sheets("6T_PRINT_1").Range("M1:AA47")
    Set prng_3_1 = ThisWorkbook.Sheets("6T_PRINT_1").Range("AB1:AP47")
    Set prng_4_1 = ThisWorkbook.Sheets("6T_PRINT_1").Range("AQ1:BE47")
    Set prng_5_1 = ThisWorkbook.Sheets("6T_PRINT_1").Range("BF1:BT47")
    Set prng_6_1 = ThisWorkbook.Sheets("6T_PRINT_1").Range("BU1:CI47")

    Set srng_1_2 = ThisWorkbook.Sheets("Scope_Change").Range("A48:L94")
    Set srng_2_2 = ThisWorkbook.Sheets("Scope_Change").Range("M48:AA94")
    Set srng_3_2 = ThisWorkbook.Sheets("Scope_Change").Range("AB48:AP94")
    Set srng_4_2 = ThisWorkbook.Sheets("Scope_Change").Range("AQ48:BE94")
    Set srng_5_2 = ThisWorkbook.Sheets("Scope_Change").Range("BF48:BT94")
    Set srng_6_2 = ThisWorkbook.Sheets("Scope_Change").Range("BU48:CI94")

    Set prng_1_2 = ThisWorkbook.Sheets("6T_PRINT_2").Range("A1:L47")
    Set prng_2_2 = ThisWorkbook.Sheets("6T_PRINT_2").Range("M1:AA47")
    Set prng_3_2 = ThisWorkbook.Sheets("6T_PRINT_2").Range("AB1:AP47")
    Set prng_4_2 = ThisWorkbook.Sheets("6T_PRINT_2").Range("AQ1:BE47")
    Set prng_5_2 = ThisWorkbook.Sheets("6T_PRINT_2").Range("BF1:BT47")
    Set prng_6_2 = ThisWorkbook.Sheets("6T_PRINT_2").Range("BU1:CI47")

    Set rerng_1_1 = ThisWorkbook.Sheets("Rework").Range("A1:L47")
    Set rerng_2_1 = ThisWorkbook.Sheets("Rework").Range("M1:AA47")
    Set rerng_3_1 = ThisWorkbook.Sheets("Rework").Range("AB1:AP47")
    Set rerng_4_1 = ThisWorkbook.Sheets("Rework").Range("AQ1:BE47")
    Set rerng_5_1 = ThisWorkbook.Sheets("Rework").Range("BF1:BT47")
    Set rerng_6_1 = ThisWorkbook.Sheets("Rework").Range("BU1:CI47")

    Set reprng_1_1 = ThisWorkbook.Sheets("7T_PRINT_1").Range("A1:L47")
    Set reprng_2_1 = ThisWorkbook.Sheets("7T_PRINT_1").Range("M1:AA47")
    Set reprng_3_1 = ThisWorkbook.Sheets("7T_PRINT_1").Range("AB1:AP47")
    Set reprng_4_1 = ThisWorkbook.Sheets("7T_PRINT_1").Range("AQ1:BE47")
    Set reprng_5_1 = ThisWorkbook.Sheets("7T_PRINT_1").Range("BF1:BT47")
    Set reprng_6_1 = ThisWorkbook.Sheets("7T_PRINT_1").Range("BU1:CI47")

    Set rerng_1_2 = ThisWorkbook.Sheets("Rework").Range("A48:L94")
    Set rerng_2_2 = ThisWorkbook.Sheets("Rework").Range("M48:AA94")
    Set rerng_3_2 = ThisWorkbook.Sheets("Rework").Range("AB48:AP94")
    Set rerng_4_2 = ThisWorkbook.Sheets("Rework").Range("AQ48:BE94")
    Set rerng_5_2 = ThisWorkbook.Sheets("Rework").Range("BF48:BT94")
    Set rerng_6_2 = ThisWorkbook.Sheets("Rework").Range("BU48:CI94")

    Set reprng_1_2 = ThisWorkbook.Sheets("7T_PRINT_2").Range("A1:L47")
    Set reprng_2_2 = ThisWorkbook.Sheets("7T_PRINT_2").Range("M1:AA47")
    Set reprng_3_2 = ThisWorkbook.Sheets("7T_PRINT_2").Range("AB1:AP47")
    Set reprng_4_2 = ThisWorkbook.Sheets("7T_PRINT_2").Range("AQ1:BE47")
    Set reprng_5_2 = ThisWorkbook.Sheets("7T_PRINT_2").Range("BF1:BT47")
    Set reprng_6_2 = ThisWorkbook.Sheets("7T_PRINT_2").Range("BU1:CI47")

    If ThisWorkbook.Sheets("TERMINAL").Cells(1, 9).Value > 0 Then
        Application.CutCopyMode = False
        brng_1.Copy
        bprng_1.PasteSpecial xlPasteAll
    End If
    If ThisWorkbook.Sheets("TERMINAL").Cells(1, 17).Value > 0 Then
        Application.CutCopyMode = False
        brng_2.Copy
        bprng_2.PasteSpecial xlPasteAll
    End If

    If ThisWorkbook.Sheets("6T_TERMINAL_1").Cells(1, 5).Value > 0 Then
        Application.CutCopyMode = False
        srng_1_1.Copy
        prng_1_1.PasteSpecial xlPasteAll
    End If
    If ThisWorkbook.Sheets("6T_TERMINAL_2").Cells(1, 5).Value > 0 Then
        Application.CutCopyMode = False
        srng_1_2.Copy
        prng_1_2.PasteSpecial xlPasteAll
    End If

    If ThisWorkbook.Sheets("6T_TERMINAL_1").Cells(1, 14).Value > 0 Then
        Application.CutCopyMode = False
        srng_2_1.Copy
        prng_2_1.PasteSpecial xlPasteAll
    End If
    If ThisWorkbook.Sheets("6T_TERMINAL_2").Cells(1, 14).Value > 0 Then
        Application.CutCopyMode = False
        srng_2_2.Copy
        prng_2_2.PasteSpecial xlPasteAll
    End If

    If ThisWorkbook.Sheets("6T_TERMINAL_1").Cells(1, 29).Value > 0 Then
        Application.CutCopyMode = False
        srng_3_1.Copy
        prng_3_1.PasteSpecial xlPasteAll
    End If
    If ThisWorkbook.Sheets("6T_TERMINAL_2").Cells(1, 29).Value > 0 Then
        Application.CutCopyMode = False
        srng_3_2.Copy
        prng_3_2.PasteSpecial xlPasteAll
    End If

    If ThisWorkbook.Sheets("6T_TERMINAL_1").Cells(1, 44).Value > 0 Then
        Application.CutCopyMode = False
        srng_4_1.Copy
        prng_4_1.PasteSpecial xlPasteAll
    End If
    If ThisWorkbook.Sheets("6T_TERMINAL_2").Cells(1, 44).Value > 0 Then
        Application.CutCopyMode = False
        srng_4_2.Copy
        prng_4_2.PasteSpecial xlPasteAll
    End If

    If ThisWorkbook.Sheets("6T_TERMINAL_1").Cells(1, 59).Value > 0 Then
        Application.CutCopyMode = False
        srng_5_1.Copy
        prng_5_1.PasteSpecial xlPasteAll
    End If
    If ThisWorkbook.Sheets("6T_TERMINAL_2").Cells(1, 59).Value > 0 Then
        Application.CutCopyMode = False
        srng_5_2.Copy
        prng_5_2.PasteSpecial xlPasteAll
    End If

    If ThisWorkbook.Sheets("6T_TERMINAL_1").Cells(1, 74).Value > 0 Then
        Application.CutCopyMode = False
        srng_6_1.Copy
        prng_6_1.PasteSpecial xlPasteAll
    End If
    If ThisWorkbook.Sheets("6T_TERMINAL_2").Cells(1, 74).Value > 0 Then
        Application.CutCopyMode = False
        srng_6_2.Copy
        prng_6_2.PasteSpecial xlPasteAll
    End If

    If ThisWorkbook.Sheets("7T_TERMINAL_1").Cells(1, 5).Value > 0 Then
        Application.CutCopyMode = False
        rerng_1_1.Copy
        reprng_1_1.PasteSpecial xlPasteAll
    End If
    If ThisWorkbook.Sheets("7T_TERMINAL_2").Cells(1, 5).Value > 0 Then
        Application.CutCopyMode = False
        rerng_1_2.Copy
        reprng_1_2.PasteSpecial xlPasteAll
    End If

    If ThisWorkbook.Sheets("7T_TERMINAL_1").Cells(1, 14).Value > 0 Then
        Application.CutCopyMode = False
        rerng_2_1.Copy
        reprng_2_1.PasteSpecial xlPasteAll
    End If
    If ThisWorkbook.Sheets("7T_TERMINAL_2").Cells(1, 14).Value > 0 Then
        Application.CutCopyMode = False
        rerng_2_2.Copy
        reprng_2_2.PasteSpecial xlPasteAll
    End If

    If ThisWorkbook.Sheets("7T_TERMINAL_1").Cells(1, 29).Value > 0 Then
        Application.CutCopyMode = False
        rerng_3_1.Copy
        reprng_3_1.PasteSpecial xlPasteAll
    End If
    If ThisWorkbook.Sheets("7T_TERMINAL_2").Cells(1, 29).Value > 0 Then
        Application.CutCopyMode = False
        rerng_3_2.Copy
        reprng_3_2.PasteSpecial xlPasteAll
    End If

    If ThisWorkbook.Sheets("7T_TERMINAL_1").Cells(1, 44).Value > 0 Then
        Application.CutCopyMode = False
        rerng_4_1.Copy
        reprng_4_1.PasteSpecial xlPasteAll
    End If
    If ThisWorkbook.Sheets("7T_TERMINAL_2").Cells(1, 44).Value > 0 Then
        Application.CutCopyMode = False
        rerng_4_2.Copy
        reprng_4_2.PasteSpecial xlPasteAll
    End If

    If ThisWorkbook.Sheets("7T_TERMINAL_1").Cells(1, 59).Value > 0 Then
        Application.CutCopyMode = False
        rerng_5_1.Copy
        reprng_5_1.PasteSpecial xlPasteAll
    End If
    If ThisWorkbook.Sheets("7T_TERMINAL_2").Cells(1, 59).Value > 0 Then
        Application.CutCopyMode = False
        rerng_5_2.Copy
        reprng_5_2.PasteSpecial xlPasteAll
    End If

    If ThisWorkbook.Sheets("7T_TERMINAL_1").Cells(1, 74).Value > 0 Then
        Application.CutCopyMode = False
        rerng_6_1.Copy
        reprng_6_1.PasteSpecial xlPasteAll
    End If
    If ThisWorkbook.Sheets("7T_TERMINAL_2").Cells(1, 74).Value > 0 Then
        Application.CutCopyMode = False
        rerng_6_2.Copy
        reprng_6_2.PasteSpecial xlPasteAll
    End If
    
    ThisWorkbook.Sheets("PRINT").ResetAllPageBreaks
    ThisWorkbook.Sheets("PRINT").VPageBreaks.Add before:=Columns("P")
    ThisWorkbook.Sheets("Project_Budget").ResetAllPageBreaks
    ThisWorkbook.Sheets("Project_Budget").VPageBreaks.Add before:=Columns("P")

    ThisWorkbook.Sheets("PRINT").Visible = xlSheetVisible
    ThisWorkbook.Sheets("6T_PRINT_1").Visible = xlSheetVisible
    ThisWorkbook.Sheets("6T_PRINT_2").Visible = xlSheetVisible
    ThisWorkbook.Sheets("7T_PRINT_1").Visible = xlSheetVisible
    ThisWorkbook.Sheets("7T_PRINT_2").Visible = xlSheetVisible
'''    ThisWorkbook.Sheets(Array("DASHBOARD", "PRINT", "6T_PRINT_1", "6T_PRINT_2","7T_PRINT_1","7T_PRINT_2")).PrintOut Copies:=1, ActivePrinter:="PDFCreator", printtofile:=True, collate:=True, prtofilename:=PSFileName
   ThisWorkbook.Sheets(Array("DASHBOARD", "PRINT", "6T_PRINT_1", "6T_PRINT_2", "7T_PRINT_1", "7T_PRINT_2")).Select
   ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ThisWorkbook.Path & "\" & "Financial Budget Design for Budget", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True

    Call Hide

    ThisWorkbook.Worksheets("Project_Budget").Activate

exitsub:
    Application.EnableEvents = True

End Sub

Sub Project_Budget_M()

    Call Protect
    
    ThisWorkbook.Sheets("Project_Budget").Range("G2").Value = ThisWorkbook.Sheets("Project_Budget").Range("A14").Value
    ThisWorkbook.Sheets("Project_Budget").Range("G3").Value = ThisWorkbook.Sheets("Project_Budget").Range("A17").Value
    ThisWorkbook.Sheets("Project_Budget").Range("I2").Value = ThisWorkbook.Sheets("Project_Budget").Range("A20").Value
    ThisWorkbook.Sheets("Project_Budget").Range("I3").Value = ThisWorkbook.Sheets("Project_Budget").Range("A23").Value
    ThisWorkbook.Sheets("Project_Budget").Range("K2").Value = ThisWorkbook.Sheets("Project_Budget").Range("A26").Value
    
    ThisWorkbook.Sheets("TERMINAL").Range("A4:A35").Value = ThisWorkbook.Sheets("Project_Budget").Range("A4:A35").Value
    ThisWorkbook.Sheets("TERMINAL").Range("B4:C30").Value = ThisWorkbook.Sheets("Project_Budget").Range("B4:C30").Value
    ThisWorkbook.Sheets("TERMINAL").Range("H4:AD30").Value = ThisWorkbook.Sheets("Project_Budget").Range("H4:AD30").Value
'''    ThisWorkbook.Sheets("TERMINAL").Range("B33:C33").Value = ThisWorkbook.Sheets("Project_Budget").Range("B33:C33").Value
    ThisWorkbook.Sheets("TERMINAL").Range("C36:C43").Value = ThisWorkbook.Sheets("Project_Budget").Range("C36:C43").Value
    ThisWorkbook.Sheets("TERMINAL").Range("G2:L3").Value = ThisWorkbook.Sheets("Project_Budget").Range("G2:L3").Value

    ThisWorkbook.Sheets("Project_Budget").Range("D7:G30").Value = ThisWorkbook.Sheets("TERMINAL").Range("D7:G30").Value
    ThisWorkbook.Sheets("Project_Budget").Range("B31:AD33").Value = ThisWorkbook.Sheets("TERMINAL").Range("B31:AD33").Value
    ThisWorkbook.Sheets("Project_Budget").Range("C44:C45").Value = ThisWorkbook.Sheets("TERMINAL").Range("C44:C45").Value
    
    ThisWorkbook.Sheets("Project_Budget").Range("H37").Value = ThisWorkbook.Sheets("Project_Budget").Range("C31").Value
    ThisWorkbook.Sheets("Project_Budget").Range("H40").Value = ThisWorkbook.Sheets("Project_Budget").Range("C32").Value

    ThisWorkbook.Sheets("DASHBOARD").Range("I3").Value = ThisWorkbook.Sheets("Project_Budget").Range("B1").Value
    ThisWorkbook.Sheets("DASHBOARD").Range("I4").Value = ThisWorkbook.Sheets("Project_Budget").Range("B2").Value
    ThisWorkbook.Sheets("DASHBOARD").Range("I5").Value = ThisWorkbook.Sheets("Project_Budget").Range("B3").Value
    
    Dim re_vert As Integer
    Dim re_verc As Integer
    
    With ThisWorkbook.Sheets("Project_Budget").Cells.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With

    For y = 7 To 30
        If ThisWorkbook.Sheets("Project_Budget").Cells(y, 7).Value < 0 Then
            re_vert = 0
            re_verc = 0
            For a = 8 To 14 Step 2
                re_vert = ThisWorkbook.Sheets("Project_Budget").Cells(y, a).Value + re_vert
                If ThisWorkbook.Sheets("Project_Budget").Cells(y, 2).Value - re_vert < 0 Then
                    With Cells(y, a).Font
                        .Color = -16776961
                        .TintAndShade = 0
                    End With
                End If
            Next a
            For b = 9 To 15 Step 2
                re_verc = ThisWorkbook.Sheets("Project_Budget").Cells(y, b).Value + re_verc
                If ThisWorkbook.Sheets("Project_Budget").Cells(y, 3).Value - re_verc < 0 Then
                    With Cells(y, b).Font
                        .Color = -16776961
                        .TintAndShade = 0
                    End With
                End If
            Next b
            For c = 17 To 29 Step 2
                re_vert = ThisWorkbook.Sheets("Project_Budget").Cells(y, c).Value + re_vert
                If ThisWorkbook.Sheets("Project_Budget").Cells(y, 2).Value - re_vert < 0 Then
                    With Cells(y, c).Font
                        .Color = -16776961
                        .TintAndShade = 0
                    End With
                End If
            Next c
            For d = 18 To 30 Step 2
                re_verc = ThisWorkbook.Sheets("Project_Budget").Cells(y, d).Value + re_verc
                If ThisWorkbook.Sheets("Project_Budget").Cells(y, 3).Value - re_verc < 0 Then
                    With Cells(y, d).Font
                        .Color = -16776961
                        .TintAndShade = 0
                    End With
                End If
            Next d
        End If
    Next y
End Sub

Sub Scope_Change_M()

    Call Protect
    
    ThisWorkbook.Sheets("Scope_Change").Range("A49:CM49").Value = ThisWorkbook.Sheets("Scope_Change").Range("A2:CM2").Value
    
    ThisWorkbook.Sheets("TERMINAL_6T").Range("E5:CM46").Value = ThisWorkbook.Sheets("Scope_Change").Range("E5:CM46").Value
    ThisWorkbook.Sheets("TERMINAL_6T").Range("E51:CM93").Value = ThisWorkbook.Sheets("Scope_Change").Range("E51:CM93").Value
    
    ThisWorkbook.Sheets("6T_TERMINAL_1").Range("E5:CM46").Value = ThisWorkbook.Sheets("Scope_Change").Range("E5:CM46").Value
    ThisWorkbook.Sheets("6T_TERMINAL_2").Range("E51:CM93").Value = ThisWorkbook.Sheets("Scope_Change").Range("E51:CM93").Value
    
    ThisWorkbook.Sheets("Scope_Change").Range("C5:D46").Value = ThisWorkbook.Sheets("TERMINAL_6T").Range("C5:D46").Value
    ThisWorkbook.Sheets("Scope_Change").Range("C51:D93").Value = ThisWorkbook.Sheets("TERMINAL_6T").Range("C51:D93").Value
    ThisWorkbook.Sheets("Scope_Change").Range("C4:CM4").Value = ThisWorkbook.Sheets("TERMINAL_6T").Range("C4:CM4").Value
    
    ThisWorkbook.Sheets("Scope_Change").Range("M1:M94").Value = ThisWorkbook.Sheets("Scope_Change").Range("A1:A94").Value
    ThisWorkbook.Sheets("Scope_Change").Range("AB1:AB94").Value = ThisWorkbook.Sheets("Scope_Change").Range("A1:A94").Value
    ThisWorkbook.Sheets("Scope_Change").Range("AQ1:AQ94").Value = ThisWorkbook.Sheets("Scope_Change").Range("A1:A94").Value
    ThisWorkbook.Sheets("Scope_Change").Range("BF1:BF94").Value = ThisWorkbook.Sheets("Scope_Change").Range("A1:A94").Value
    ThisWorkbook.Sheets("Scope_Change").Range("BU1:BU94").Value = ThisWorkbook.Sheets("Scope_Change").Range("A1:A94").Value

End Sub

Sub Rework_M()

    Call Protect
    
    ThisWorkbook.Sheets("Rework").Range("A49:CM49").Value = ThisWorkbook.Sheets("Rework").Range("A2:CM2").Value
    
    ThisWorkbook.Sheets("TERMINAL_7T").Range("E5:CM46").Value = ThisWorkbook.Sheets("Rework").Range("E5:CM46").Value
    ThisWorkbook.Sheets("TERMINAL_7T").Range("E51:CM93").Value = ThisWorkbook.Sheets("Rework").Range("E51:CM93").Value
    
    ThisWorkbook.Sheets("7T_TERMINAL_1").Range("E5:CM46").Value = ThisWorkbook.Sheets("Rework").Range("E5:CM46").Value
    ThisWorkbook.Sheets("7T_TERMINAL_2").Range("E51:CM93").Value = ThisWorkbook.Sheets("Rework").Range("E51:CM93").Value
    
    ThisWorkbook.Sheets("Rework").Range("C5:D46").Value = ThisWorkbook.Sheets("TERMINAL_7T").Range("C5:D46").Value
    ThisWorkbook.Sheets("Rework").Range("C51:D93").Value = ThisWorkbook.Sheets("TERMINAL_7T").Range("C51:D93").Value
    ThisWorkbook.Sheets("Rework").Range("C4:CM4").Value = ThisWorkbook.Sheets("TERMINAL_7T").Range("C4:CM4").Value
    
    ThisWorkbook.Sheets("Rework").Range("M1:M94").Value = ThisWorkbook.Sheets("Rework").Range("A1:A94").Value
    ThisWorkbook.Sheets("Rework").Range("AB1:AB94").Value = ThisWorkbook.Sheets("Rework").Range("A1:A94").Value
    ThisWorkbook.Sheets("Rework").Range("AQ1:AQ94").Value = ThisWorkbook.Sheets("Rework").Range("A1:A94").Value
    ThisWorkbook.Sheets("Rework").Range("BF1:BF94").Value = ThisWorkbook.Sheets("Rework").Range("A1:A94").Value
    ThisWorkbook.Sheets("Rework").Range("BU1:BU94").Value = ThisWorkbook.Sheets("Rework").Range("A1:A94").Value

End Sub

Sub Protect()

    Worksheets("DASHBOARD").Protect "password", userinterfaceonly:=True
    Worksheets("Project_Budget").Protect "password", userinterfaceonly:=True
    Worksheets("Scope_Change").Protect "password", userinterfaceonly:=True
    Worksheets("Rework").Protect "password", userinterfaceonly:=True

End Sub

Sub UnProtect()

    For i = 1 To Sheets.Count
        Sheets(i).UnProtect "password"
    Next i
    
End Sub

Sub UnHide()

    ThisWorkbook.Sheets("PRINT").Visible = True
    ThisWorkbook.Sheets("TERMINAL").Visible = True
    ThisWorkbook.Sheets("6T_PRINT_1").Visible = True
    ThisWorkbook.Sheets("6T_PRINT_2").Visible = True
    ThisWorkbook.Sheets("6T_TERMINAL_1").Visible = True
    ThisWorkbook.Sheets("6T_TERMINAL_2").Visible = True
    ThisWorkbook.Sheets("TERMINAL_6T").Visible = True
    ThisWorkbook.Sheets("7T_PRINT_1").Visible = True
    ThisWorkbook.Sheets("7T_PRINT_2").Visible = True
    ThisWorkbook.Sheets("7T_TERMINAL_1").Visible = True
    ThisWorkbook.Sheets("7T_TERMINAL_2").Visible = True
    ThisWorkbook.Sheets("TERMINAL_7T").Visible = True


End Sub

Sub Hide()

    ThisWorkbook.Sheets("PRINT").Visible = xlVeryHidden
    ThisWorkbook.Sheets("TERMINAL").Visible = xlVeryHidden
    ThisWorkbook.Sheets("6T_PRINT_1").Visible = xlVeryHidden
    ThisWorkbook.Sheets("6T_PRINT_2").Visible = xlVeryHidden
    ThisWorkbook.Sheets("6T_TERMINAL_1").Visible = xlVeryHidden
    ThisWorkbook.Sheets("6T_TERMINAL_2").Visible = xlVeryHidden
    ThisWorkbook.Sheets("TERMINAL_6T").Visible = xlVeryHidden
    ThisWorkbook.Sheets("7T_PRINT_1").Visible = xlVeryHidden
    ThisWorkbook.Sheets("7T_PRINT_2").Visible = xlVeryHidden
    ThisWorkbook.Sheets("7T_TERMINAL_1").Visible = xlVeryHidden
    ThisWorkbook.Sheets("7T_TERMINAL_2").Visible = xlVeryHidden
    ThisWorkbook.Sheets("TERMINAL_7T").Visible = xlVeryHidden

End Sub
