Attribute VB_Name = "Modulo1"
Sub FormatCSV()
Attribute FormatCSV.VB_ProcData.VB_Invoke_Func = "f\n14"
'
' FormatCSV Macro
'
' Scelta rapida da tastiera: CTRL+f
'
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
        Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1)), DecimalSeparator:=".", _
        ThousandsSeparator:=",", TrailingMinusNumbers:=True
End Sub
