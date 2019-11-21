Attribute VB_Name = "Module2"
Sub WorksheetLoop()

'Displays options for user.  See string in message box below for explanation of options.
What_To_Do = MsgBox("Press Yes to run stocks macro on all worksheets, No to run macro only on the active sheet, and Cancel to exit", vbYesNoCancel)

If What_To_Do = vbYes Then
    'Original Source for this part in link below.
    'https://support.microsoft.com/en-us/help/142126/macro-to-loop-through-all-worksheets-in-a-workbook
    
    Dim WS_Count As Integer
    Dim I As Integer
    

    
    ' Set WS_Count equal to the number of worksheets in the active
    ' workbook.
    WS_Count = ActiveWorkbook.Worksheets.Count

    ' Begin the loop.
    For I = 1 To WS_Count
        
        Worksheets(I).Activate  'Added line to change active worksheet
        Call stocks 'Runs Sub stocks on the active worksheet
        ' Insert your code here.
        ' The following line shows how to reference a sheet within
        ' the loop by displaying the worksheet name in a dialog box.
        'MsgBox ActiveWorkbook.Worksheets(I).Name

    Next I
ElseIf What_To_Do = vbNo Then
    Call stocks 'Runs Sub stocks on the active worksheet
Else
    MsgBox ("No changes made.")  'Exits without doing anything
End If


End Sub
