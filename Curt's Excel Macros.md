---
layout: page
title: Excel Macros
---

```VBnet
Sub Format_PTFields()
'Macro goal: allow users to quickly choose the format to apply to pivot table fields
'No keyboard shortcut assigned
'Code modified from Dick Kusleika's code at:
'http://www.dailydoseofexcel.com/archives/2010/06/18/formatting-pivot-tables/

    Dim pf As PivotField
    Dim FormatChoice As String 'allows you to dynamically select the format
   Dim QuestionString As String

    On Error GoTo HandleErr

    If TypeName(Selection) = "Range" Then Set pf = ActiveCell.PivotField


    'Consolidates the question blurb to a variable
   QuestionString = "Apply which format to this pivot field?" & vbCrLf & _
                "    '0': numbers with 0 digits after the decimals" & vbCrLf & _
                "    '1': numbers with 1 digit after the decimals" & vbCrLf & _
                "    'd': dollars (no cents)" & vbCrLf & _
                "    'c': dollars and cents"

    'Ask the user what format to apply
   FormatChoice = InputBox(QuestionString)

    'based on the FormatChoice, format the selected pivot field
   Select Case FormatChoice
        Case 0      'shows numbers with 0 digits after the decimal
           pf.NumberFormat = "#,##0"

        Case 1      'shows numbers with 1 digit after the decimal
          pf.NumberFormat = "#,##0.0"

        Case "d"    'shows dollars (no cents)
           pf.NumberFormat = "$#,##0"

        Case "c"    'Shows dollars and cents
           pf.NumberFormat = "$#,##0.00"
    End Select

ExitSub:
    Exit Sub

HandleErr:
    If Err = 1004 Then
        MsgBox ("This macro only does something useful if you are " & vbCrLf & _
                "in a pivot table value field.  Exiting macro.")
    Else
        MsgBox "Unexpected Error: " & Err & Err.Description
    End If

    GoTo ExitSub

End Sub
```

```VBnet

Sub SelectAdjacentCol()
'Selects all empty rows of adjacent column. Handy when there's no endpoint to a column in order to do an easy autofill
' Keyboard Shortcut: Ctrl+m
'Application.OnKey "^m", "SelectAdjacentCol"
    Dim rAdjacent As Range

    If TypeName(Selection) = "Range" Then
        If Selection.Column > 1 Then
            If Selection.Cells.Count = 1 Then
                If Not IsEmpty(Selection.Offset(0, -1).Value) Then
                    With Selection.Offset(0, -1)
                        Set rAdjacent = .Parent.Range(.Cells(1), .End(xlDown))
                    End With

                    Selection.Resize(rAdjacent.Cells.Count).Select
                End If
            End If
        End If
    End If
End Sub

```

```VBnet

Sub Fill_Blank_Cells()
'Fills all blank cells in the whole range. Can't exactly remember why this is useful now. Shrug
'No keyboard shortcut assigned
Selection.SpecialCells(xlCellTypeBlanks).Select
Selection.FormulaR1C1 = "=R[-1]C"
End Sub

```

```VBnet

Sub format()
' format Macro - sets font and size to be same for whole sheet
' Keyboard Shortcut: Ctrl+w
'Application.OnKey "^w", "format"
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    With Selection.Font
        .Name = "Tahoma"
        .Size = 11
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
    With Selection.Font
        .Name = "Tahoma"
        .Size = 10
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
End Sub

```

```VBnet

Sub Adjust_cols()
' Adjust_cols Macro - sets column to size of max text
' Keyboard Shortcut: Ctrl+j
'Application.OnKey "^j", "Adjust_cols"
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    ActiveCell.Columns("A:A").EntireColumn.EntireColumn.AutoFit
    Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Columns.AutoFit
    Range("A1").Select
End Sub

```

```VBnet
Sub Header()
' Header Macro - sets header row to slighter larger font. White font on black background.
' Keyboard Shortcut: Ctrl+h
'Application.OnKey "^h", "Header"
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    With Selection.Font
        .Name = "Tahoma"
        .Size = 11
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
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Columns.AutoFit
    Selection.End(xlToLeft).Select

End Sub

```

```VBnet

Sub delete_sheet()
' deletes current sheet
    ActiveWindow.SelectedSheets.Delete
' Keyboard Shortcut: Ctrl+g
    'Application.OnKey "^g", "delete_sheet"
End Sub

```

```VBnet

Sub Clear_Range_End()
' Clear_Range_End Macro - finds end of range and resets it to current last row of actual data.
' Keyboard Shortcut: Ctrl+k
'Application.OnKey "^k", "Clear_Range_End"
    ActiveSheet.UsedRange
    ActiveWorkbook.Save
End Sub
```

### This one needs to be placed in [personal.xslb] VBA [This Workbook] and change the letters to whatever may suit you best

![VBA](https://github.com/curtlymartin/Excel/blob/master/11_43_15_Microsoft%20Visual%20Basic%20for%20Applications%20-%20PERSONAL.XLSB%20-%20%5BThisWorkbook%20(Code)%5D_000119.png "Where to paste in personal.xslb VBA")

```VBnet
Private Sub workbook_open()
    Application.OnKey "^m", "SelectAdjacentCol"
    Application.OnKey "^k", "Clear_Range_End"
    Application.OnKey "^h", "Header"
    Application.OnKey "^j", "Adjust_cols"
    Application.OnKey "^w", "format"
    Application.OnKey "^g", "delete_sheet"
    Application.OnKey "^l", "Format_PTFields"
End Sub
```


.
