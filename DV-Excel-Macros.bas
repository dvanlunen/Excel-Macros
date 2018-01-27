Sub TableFormat()
  ' Starting in the top left of a table
  '    run this macro to format the table



  Dim Start As Range
  Dim TopRow As Range
  Dim TableBody As Range

  'Identify where the table is
  Set Start = Selection
  Set TopRow = Range(Start, Start.End(xlToRight))
  Set TableBody = Range(Start.Offset(1, 0), ActiveCell.SpecialCells(xlLastCell))

  'Change Font and Size of the whole sheet
  With Cells.Font
      .Name = "Arial"
      .Size = 10
  End With

  'Change the formatting of the Top Row
  TopRow.Font.Bold = True
  With TopRow.Borders(xlEdgeTop)
      .LineStyle = xlContinuous
      .ColorIndex = 0
      .TintAndShade = 0
      .Weight = xlThin
  End With
  With TopRow.Borders(xlEdgeBottom)
      .LineStyle = xlContinuous
      .ColorIndex = 0
      .TintAndShade = 0
      .Weight = xlMedium
  End With
  With TopRow
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlCenter
      .WrapText = True
  End With

  'Change the formatting of the body of the table
  With TableBody.Borders(xlEdgeTop)
      .LineStyle = xlContinuous
      .ColorIndex = 0
      .TintAndShade = 0
      .Weight = xlMedium
  End With
  With TableBody.Borders(xlEdgeBottom)
      .LineStyle = xlDouble
      .ColorIndex = 0
      .TintAndShade = 0
      .Weight = xlThick
  End With
  With TableBody.Borders(xlInsideHorizontal)
      .LineStyle = xlContinuous
      .ColorIndex = 0
      .TintAndShade = 0
      .Weight = xlHairline
  End With


  'Center all cells
  Cells.VerticalAlignment = xlCenter

  'Add filters to the headers
  ActiveSheet.AutoFilterMode = False
  Range(Start, ActiveCell.SpecialCells(xlLastCell)).AutoFilter

  'Add rows above the table
  Start.EntireRow.Insert
  Start.EntireRow.Insert
  Start.EntireRow.Insert

  'Autofit the column width
  Cells.EntireColumn.AutoFit


  'Add Freezepanes so the Header sticks on scroll
  Start.Offset(1, 0).Select
  ActiveWindow.FreezePanes = True

  'Add text to indicate a title can be added
  Start.Offset(-3, 0).Value = "Title"
  With Start.Offset(-3, 0).Font
      .Name = "Arial"
      .Size = 12
      .Bold = True
  End With

  'Change the zoom and get rid of gridlines
  ActiveWindow.Zoom = 80
  ActiveWindow.DisplayGridlines = False
End Sub




'This is a workaround to get the Macro with parameters to appear in the list
Sub BarsUnder
  Call BarsUnderSub
End Sub

Sub BarsUnderSub(Optional Cols2Compare As Integer = -1)

  '
  ' This macro is used with datasets the begin with a few columns that identify groups
  ' The user identifies how many columns define a group
  ' Then a red line is placed under the end of each group
  '

  Dim Start As Range
  Dim FirstCol As Range
  Dim AddedCol As Range
  Dim TableBody As Range
  Dim myValue As Variant
  Dim NumberAsString As String
  Dim cond As String
  Dim Forml As String
  Dim I As Integer

  ' If -1, then need user to specify number of columns to use
  If Cols2Compare = -1 Then
     'Ask for user input
     NumberAsString = InputBox(Prompt:="Enter the number of columns to compare", _
     Title:="Number needed", Default:="Enter number here")

     'Validate user input
     If Not IsNumeric(NumberAsString) Then
       MsgBox "You can only enter a number in this field"
       Exit Sub
     ElseIf Val(NumberAsString) < 1 Or Round(Val(NumberAsString)) <> Val(NumberAsString) Then
       MsgBox "You can only enter a positive number of columns"
     Else
       Cols2Compare = Val(NumberAsString)
     End If
  End If

  'Identify where the table is
  Set Start = Selection
  Set FirstCol = Range(Start, Start.End(xlDown))
  Set TableBody = Range(Start, ActiveCell.SpecialCells(xlLastCell))
  Start.EntireColumn.Insert
  Set AddedCol = FirstCol.Offset(0, -1)

  'Build condition string AND statement
  cond = "RC[1]=R[1]C[1]"
  If Cols2Compare > 1 Then
    For I = 2 To Cols2Compare
       cond = cond & ",RC[" & _
       CStr(I) & _
       "]=R[1]C[" & _
       CStr(I) & _
       "]"
    Next I
  End If

  'Apply the condition to the added row
  AddedCol.FormulaR1C1 = "=IF(AND(" & cond & "),0,1)"

  'Set up the formula for conditional formatting
  Forml = "=" & Start.Offset(0, -1).Address(False, True) & "=1"

  'Conditionally Format
  TableBody.FormatConditions.Add Type:=xlExpression, Formula1:=Forml
  TableBody.FormatConditions(TableBody.FormatConditions.Count).SetFirstPriority
  With TableBody.FormatConditions(1).Borders(xlBottom)
     .LineStyle = xlContinuous
     .Color = -16776961
  End With
  TableBody.FormatConditions(1).StopIfTrue = False



  'Hide the added column
  AddedCol.EntireColumn.Font.Color = RGB(255, 255, 255)
  AddedCol.EntireColumn.Hidden = True

End Sub





Sub WorksheetLoop()
     ' Structure from ms support
     ' This code loops through Worksheets
     ' currently it applies the TableFormat Macro to all worksheets
     ' but it can be used for a variety of purposes

     ' Declare Current as a worksheet object variable.
     Dim ws As Worksheet

     ' Loop through all of the worksheets in the active workbook.
     For Each ws In Worksheets
        'Delete the below and make your own code
        ws.Activate
        ws.Range("A1").Select
        Call TableFormat
     Next
End Sub




'This is a workaround to get the Macro with parameters to appear in the list
Sub AlternateRowShade
  Call AlternateRowShadeSub
End Sub

Sub AlternateRowShadeSub(Optional Cols2Compare As Integer = -1)
  '
  ' This macro is used with datasets the begin with a few columns that identify groups
  ' The user identifies how many columns define a group
  ' Then groups are alternatingly highlighted to help visually separate them
  '

  Dim Start As Range
  Dim FirstCol As Range
  Dim AddedCol As Range
  Dim TableBody As Range
  Dim myValue As Variant
  Dim NumberAsString As String
  Dim cond As String
  Dim Forml As String
  Dim I As Integer


  ' If -1, then need user to specify number of columns to use
  If Cols2Compare = -1 Then
      'Ask for user input
      NumberAsString = InputBox(Prompt:="Enter the number of columns to compare", _
      Title:="Number needed", Default:="Enter number here")
      'Validate user input
      If Not IsNumeric(NumberAsString) Then
        MsgBox "You can only enter a number in this field"
        Exit Sub
      ElseIf Val(NumberAsString) < 1 Or Round(Val(NumberAsString)) <> Val(NumberAsString) Then
        MsgBox "You can only enter a positive number of columns"
      Else
        Cols2Compare = Val(NumberAsString)
      End If
  End If

  'Identify where the table is
  Set Start = Selection
  Set FirstCol = Range(Start, Start.End(xlDown))
  Set TableBody = Range(Start, ActiveCell.SpecialCells(xlLastCell))

  'Add a column to hold values that will be used in the conditional formatting
  Start.EntireColumn.Insert
  Set AddedCol = FirstCol.Offset(0, -1)

  'initial value that the other values we be built off of
  Start.Offset(-1, -1).Value = "1"

  'Build condition string AND statement
  cond = "RC[1]=R[-1]C[1]"
  If Cols2Compare > 1 Then
    For I = 2 To Cols2Compare
        cond = cond & ",RC[" & _
        CStr(I) & _
        "]=R[-1]C[" & _
        CStr(I) & _
        "]"
    Next I
  End If

  'Apply the condition to the added row
  AddedCol.FormulaR1C1 = "=IF(AND(" & cond & "),R[-1]C,-1*R[-1]C)"

  'Set up the formula for coniditional formatting
  Forml = "=" & Start.Offset(0, -1).Address(False, True) & "=1"

  'Conditionally Format
  TableBody.FormatConditions.Add Type:=xlExpression, Formula1:=Forml
  With TableBody.FormatConditions(1).Interior
      .PatternColorIndex = xlAutomatic
      .ThemeColor = xlThemeColorAccent1
      .TintAndShade = 0.799981688894314
  End With
  TableBody.FormatConditions(1).StopIfTrue = False

  'Hide the added column
  AddedCol.EntireColumn.Font.Color = RGB(255, 255, 255)
  AddedCol.EntireColumn.Hidden = True

End Sub





'Format dollars rounded to the nearest dollar with commas
Sub CurrencyFormat()
    Selection.NumberFormat = "$#,##0"
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
End Sub

'Format numbers rounded to the integer with commas
Sub CommaFormat()
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    Selection.NumberFormat = "#,##0"
End Sub




Sub NotesSources()
    '
    ' Practice good documentation hygiene
    ' Add Notes and Sources
    '

    Dim Start As Range

    'Reference location
    Set Start = Selection

    'Add formatted text
    Start.Font.Bold = True
    Start.FormulaR1C1 = "Notes:"
    Start.Offset(4, 0).FormulaR1C1 = "Sources:"
    Start.Offset(4, 0).Font.Bold = True

    'Add an extra column if necessary
    If Start.Column = 1 Then
        Start.EntireColumn.Insert
        Start.Offset(0, -1).EntireColumn.ColumnWidth = 3.5
    End If

    ' Add formatted numbers that reference each other
    ' References make for easy shifting of order
    Start.Offset(1, -1).FormulaR1C1 = "1"
    Start.Offset(2, -1).FormulaR1C1 = "=R[-1]C+1"
    Start.Offset(5, -1).FormulaR1C1 = "1"
    Start.Offset(6, -1).FormulaR1C1 = "=R[-1]C+1"

    ' More formatting
    With Range(Start.Offset(1, -1), Start.Offset(7, -1))
        .NumberFormat = """[""0""]"""
        .Font.FontStyle = "Bold"
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlRight
    End With

    With Range(Start, Start.Offset(7, 0))
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
    End With

End Sub
