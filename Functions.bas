Attribute VB_Name = "Functions"
Option Explicit

Function TabExists(r As String, Optional wbk As Workbook) As Boolean
    
    If wbk Is Nothing Then Set wbk = ActiveWorkbook
    
    TabExists = False
    
    On Error Resume Next
    TabExists = wbk.Sheets(r).Index > 0
    
End Function

Function FindRow(to_find As Range, where_find As Range) As Integer

    FindRow = 0
    
    On Error Resume Next
    
    FindRow = where_find.Find(to_find.Value, , xlValues, xlWhole).row
    
    
End Function

Function FindColumn(to_find As Range, where_find As Range) As String

    FindColumn = 0
    
    On Error Resume Next
    
    FindColumn = Split(where_find.Find(to_find.Value, , xlValues, xlWhole).Address, "$")(1)
    
End Function

Function NumberToVlookup(first As Range, secound As Range) As Integer
    
End Function

Function Letter2Number(col As String) As Integer

    Letter2Number = Split(Range(col & ":1").Address, "$")(2)
    
End Function

