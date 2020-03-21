Attribute VB_Name = "Macros"
Option Explicit

Private wbk_template As Workbook
Private wks_macrotab As Worksheet

Sub ReadyForPR()

    Application.Calculation = True
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim FDialog As FileDialog
    Set FDialog = Application.FileDialog(msoFileDialogFilePicker)
    Dim wbk_client As Workbook
    
    Dim message As Variant
    
    Set wbk_template = ThisWorkbook
    Set wks_macrotab = wbk_template.Sheets("MacroTab")
    message = ""
    
    On Error Resume Next
    
    With FDialog
        .Title = "Select our template recived from client!"
        .InitialFileName = wbk_template.Path & "\"
        .Show
    End With
    
    If Err.Number <> 0 Then
        message = "No file choosen! Check will be done on current data." & vbCrLf
        Err.Clear
    Else
        

        Workbooks.Open (FDialog.SelectedItems(1))
        Set wbk_client = Workbooks(Dir(FDialog.SelectedItems(1)))
        
        On Error GoTo Sheet_error
        wbk_client.Sheets("1 Client Info").Select
        Range("C3:C32").Copy
        
        wbk_template.Activate
        Sheets("1 Client Info").Select
        Range("C3").PasteSpecial xlPasteValues
        Range("C35").Calculate
        
        wbk_client.Close
    End If
    
    On Error GoTo 0
    
    wks_macrotab.Range("A1").Calculate
    
    If wks_macrotab.Range("A1").Value <> "Ready - info complete" Then
        message = message & vbCrLf & "Further preparation should be consulted with FR!!" & vbCrLf & vbCrLf & "Do you want run filling macro?"
        
        If MsgBox(message, vbYesNo + vbCritical, "Notification") = vbYes Then
            Call Client_Data
        End If
    Else
        message = message & vbCrLf & "You can performe your work" & vbCrLf & vbCrLf & "Do you want run filling macro?"
        
        If MsgBox(message, vbYesNo + vbQuestion, "Notification") = vbYes Then
            Call Client_Data
        End If
    End If
    
    
    
    Exit Sub
Sheet_error:
    MsgBox "Tab '1 Client Info' does not exist! ", vbCritical
        
End Sub

Sub TabColorChanging()
    Dim rng As Range
    Dim i As Integer
    
    
    Set wbk_template = ThisWorkbook
    Set wks_macrotab = wbk_template.Sheets("MacroTab")

    On Error GoTo Client_error
    Set rng = wks_macrotab.Range(wks_macrotab.Range("A6").Value)
    
    'On Error GoTo Sheet_error
    i = 0
    Do While rng.Offset(i, 1).Value <> ""
        wbk_template.Sheets(rng.Offset(i, 1).Value).Tab.Color = RGB(191, 191, 191)
        i = i + 1
    Loop
    
    Exit Sub
Client_error:
    MsgBox "Please check if 'SII QRT Type' in the '1 Client Info' tab has been chosen from drop-down list.", , "Error!"
Sheet_error:
    MsgBox "Tab " & rng.Offset(i, 1).Value & " does not exist!", vbCritical, "Error!"
    
End Sub
Sub TabColorChanging2()
    Dim rng As Range
    Dim i As Integer
    
    
    Set wbk_template = ThisWorkbook
    Set wks_macrotab = wbk_template.Sheets("MacroTab")

    On Error GoTo Client_error
    Set rng = wks_macrotab.Range("K6")
    
    wks_macrotab.Range("L6:L14").Calculate
    
    On Error GoTo Sheet_error
    i = 0
    Do While rng.Offset(i, 0).Value <> ""
        If rng.Offset(i, 1).Value = 0 Then
            wbk_template.Sheets(rng.Offset(i, 0).Value).Tab.Color = RGB(191, 191, 191)
        Else
            wbk_template.Sheets(rng.Offset(i, 0).Value).Tab.Color = RGB(146, 208, 80)
        End If
        i = i + 1
    Loop
    
    Exit Sub
Client_error:
    MsgBox "Please check if 'SII QRT Type' in the '1 Client Info' tab has been chosen from drop-down list.", , "Error!"
Sheet_error:
    MsgBox "Tab " & rng.Offset(i, 0).Value & " does not exist!", vbCritical, "Error!"
    
End Sub

Sub Client_Data()

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    
    Dim FDialog As FileDialog
    Dim wbk_client As Workbook
    Dim wks_mapping As Worksheet
    Dim wks_hiden As Worksheet
    Dim client_name As String
    Dim tb As String
    
    Dim row As Integer
    
    Set wbk_template = ThisWorkbook
    
    On Error GoTo Macros_tab_error
        Set wks_mapping = wbk_template.Sheets("Macros")
    On Error GoTo 0
    
    Set wks_hiden = wbk_template.Sheets("MacroTab")
    
    If wks_hiden.Range("E1").Value <> "" Then
    
        On Error Resume Next
            Set wbk_client = Workbooks(wks_hiden.Range("E1").Value)
            
            If Err.Number <> 0 Then
                Set wbk_client = wbk_open
            End If
        On Error GoTo 0
        
        row = 0
        
        Do While wks_mapping.Range("H6").Offset(row, 0).Value <> ""
            tb = wks_mapping.Range("I6").Offset(row, 0).Value
            If tb <> "" And wks_mapping.Range("H6").Offset(row, 0).Value <> "TRUE" Then
                If TabExists(tb, wbk_client) Then
                    wbk_client.Sheets(tb).Copy After:=wbk_template.Sheets(wbk_template.Sheets.Count)
                    ActiveSheet.Name = wks_mapping.Range("I6").Offset(row, -2).Value & "X"
                    
                End If
            End If
            row = row + 1
        Loop
        wks_hiden.Range("E1").Clear
        Call TabColorChanging2

    Else
        Set wbk_client = wbk_open
        
        
        wks_hiden.Range("E1").Value = wbk_client.Name
        Call sugestion(wbk_client)
        
    End If
    
    wks_mapping.Activate
    
    Application.Calculation = xlCalculationAutomatic
    
Exit Sub
Macros_tab_error:
    MsgBox "Tab with macros should be named as 'Macros'." & vbCrLf & "Please verify if this statement is fulfilled."
    
End Sub

Function wbk_open() As Workbook

    Dim FDialog As FileDialog
    
    Set FDialog = Application.FileDialog(msoFileDialogFilePicker)
        
    With FDialog
        .Title = "Select our template recived from client!"
        .InitialFileName = wbk_template.Path & "\"
        If .Show <> -1 Then GoTo No_file_Choosen
    End With
        
    Workbooks.Open (FDialog.SelectedItems(1))
    Set wbk_open = Workbooks(Dir(FDialog.SelectedItems(1)))
    
    Exit Function
No_file_Choosen:
    MsgBox "You have to choose client's file.", vbCritical
       
End Function

Sub sugestion(wk As Workbook)
    Dim rx As New RegExp
    Dim patter As Variant
    Dim ws As Worksheet
    Dim i As Integer
    patter = ThisWorkbook.Sheets("MacroTab").Range("O6:O15").Value
    
    For i = 1 To 10
        rx.Pattern = CStr(patter(i, 1))
        For Each ws In wk.Sheets
            If rx.Test(ws.Name) Then
                 ThisWorkbook.Sheets("MacroTab").Range("P" & (i + 5)).Value = ws.Name
            End If
        Next ws
    Next i
    
End Sub


Sub free_of_formulas()

    Application.ScreenUpdating = False
    Application.Calculation = False
    
    Dim wks_mapping As Worksheet
    Dim s_range As String
    Dim i As Integer
    
    Set wbk_template = ThisWorkbook
    Set wks_mapping = wbk_template.Sheets("MacroTab")
    
    wks_mapping.Range("L6:L15").Value = wks_mapping.Range("L6:L15").Value
    
    With wks_mapping
        For i = 6 To 15
            If (.Range("L" & i).Value = 0) Then
                wbk_template.Sheets(.Range("K" & i).Value).Range(.Range("M" & i).Value & "," & .Range("N" & i).Value).ClearContents
                
            Else
                wbk_template.Sheets(.Range("K" & i).Value).Range(.Range("M" & i).Value).Copy
                
                wbk_template.Sheets(.Range("K" & i).Value).Range(.Range("M" & i).Value).PasteSpecial xlPasteValues
                
                Call conv_to_nr(wbk_template.Sheets(.Range("K" & i).Value).Range(.Range("M" & i).Value))
                
                wbk_template.Sheets(.Range("K" & i).Value).Range(.Range("N" & i).Value).ClearContents
            End If
        Next i
    End With
    
    Call tabs_delate
    wks_macrotab.Activate
    ThisWorkbook.Activate
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub tabs_delate()

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Set wbk_template = ThisWorkbook
    Set wks_macrotab = wbk_template.Sheets("MacroTab")
    Dim i As Integer
    Dim rng As Range
    
    i = 0
    Set rng = wks_macrotab.Range("K6")
    
    Do While rng.Offset(i, 0).Value <> 0
        If TabExists(rng.Offset(i, 0).Value & "X", wbk_template) Then
            wbk_template.Sheets(rng.Offset(i, 0).Value & "X").Delete
        End If
    i = i + 1
    Loop
    
    wks_macrotab.Activate
End Sub

Sub Finalize()

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    On Error Resume Next
    
    ThisWorkbook.Sheets("MacroTab").Delete
    ThisWorkbook.Sheets("Macros").Delete
    
    ThisWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\" & "QRT Tie out template Resulta", FileFormat:=xlWorkbookDefault
    
End Sub
Sub conv_to_nr(r As Range)
    With r
        .NumberFormat = "General"
        .Value = .Value
        
    End With
End Sub


Sub pass_delate()

    Dim Validation_tab As Worksheet
    Dim wks_mapping As Worksheet
    Dim dict As New Dictionary
    Dim rx As New RegExp
    Dim r As String
    Dim k As Integer
    Dim rng As Range
    Dim i As Integer
    
    Set wks_mapping = ThisWorkbook.Sheets("MacroTab")
    
    For i = 6 To 15
        dict.Add wks_mapping.Range("K" & i).Value, wks_mapping.Range("L" & i).Value
    Next i
    
    With rx
        .IgnoreCase = True
        .Pattern = "((pass|true)|^0$)"
    End With
    
    On Error GoTo TabExist:
        Set Validation_tab = ThisWorkbook.Sheets("Validation fails")
    On Error GoTo 0
    
    Set rng = Validation_tab.Range("G5")
    i = 0
    k = 5
    Do While rng.Offset(i, -5).Value <> 0
        If Not (rx.Test(CStr(rng.Offset(i, 0).Value)) Or dict.Item(rng.Offset(i, -4).Value) = 0) Then
            If k = i + 5 Then
                k = k + 1
            Else
                r = r & k & ":" & (i + 4) & ","
                k = i + 6
            End If
        End If
        i = i + 1
    Loop
    
    If k < i + 5 Then r = r & k & ":" & (i + 4) & ","
    
    On Error Resume Next
    Validation_tab.Range(Left(r, Len(r) - 1)).Delete
    On Error GoTo 0
    
    Validation_tab.Range("B5").Value = "1"
    Validation_tab.Range("B6").Value = "2"
    Validation_tab.Range("B5:B6").AutoFill Destination:=Range(Validation_tab.Range("B5"), Validation_tab.Range("B5").End(xlDown))
Exit Sub
TabExist:
    MsgBox "'Validation fails' tab does not exists"
End Sub

