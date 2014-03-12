Attribute VB_Name = "Copy_Sheet_Insert"
Sub CpShtIns() 'CpShtIns (Copy Sheet Insert) will creat a copy of the sheet specified in this macro and place it directly after the copied sheet

    Dim start_time, end_time
    start_time = Now()

    Dim shtNw As Worksheet
    Dim shtCp As String, _
    shtNm As String, _
    dataInd As String, _
    partNm As String

    Dim shtCnt As Integer
    
    '**CHANGE THIS TO CHANGE THE SHEET TO COPY**
    shtCp = "Partner_Template" 'String name of the sheet to be copied
    dataInd = "P_" 'String added to beginning of sheet name (in this case to indicate data sheet)
    
    'Defines input box for partner name, string at the front is for identification in ConsData() method
    partNm = Application.InputBox("Enter Partner Name" & vbNewLine _
        & "- 29 or less characters (assuming a 2-char prefix)" & vbNewLine _
        & "- Cannot include the following characters: / \ * ? [ ]", Type:=2)
    
    shtNm = dataInd & partNm 'Creates sheet name from input and indicator prefix
    
    If SheetExist(shtNm) = True Then
        MsgBox ("Sheet already exists, please choose a different name.")
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    If partNm <> "" And partNm <> "False" Then 'Does nothing if no value passed into input box
    
        Sheets(shtCp).Copy After:=Sheets(Sheets.Count) 'Copies sheet to end of workbook
        Set shtNw = Sheets(Sheets.Count) 'Returns newly created sheet by referencing sheet index at end of workbook
        shtNw.Name = shtNm 'Sets sheet name to variable
        shtNw.Cells(2, 2).value = partNm 'inputs partner name in sheet key
    
    End If
    
    Application.ScreenUpdating = True
    end_time = Now()
    
    MsgBox ("Created sheet: " & shtNm & vbNewLine & "Time elapsed: " & DateDiff("s", start_time, end_time) & " seconds")

End Sub

Function SheetExist(strSheetName As String) As Boolean
    Dim i As Integer

    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = strSheetName Then
            SheetExist = True
            Exit Function
        End If
    Next i
End Function
