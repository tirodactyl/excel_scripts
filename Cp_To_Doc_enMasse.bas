Attribute VB_Name = "Cp_To_ADD"
Sub cpToADD() 'copies data to a sheet in another workbook

    Dim wrkSrc As Workbook: Set wrkSrc = ThisWorkbook
    Dim shtSrc As Worksheet: Set shtSrc = ThisWorkbook.Sheets("Legacy Update")
    Dim wrkDest As Workbook 'destination workbook opened to paste data into
    Dim wrkPath As String 'path of destination workbook including file name - needed in order to open the file (MUST BE IN THE SAME PARENT FOLDER)
    Dim shtDest As Worksheet 'destination sheet to copy data into
    Dim shtDestNm As String 'string representing the name of the destination worksheet
    
    wrkPath = ":" & "Acq Deep Dive.xlsx" 'path within parent directory to the destination workbook ( ":" instead of "/")
    shtDestNm = "Total Spend Summary" 'name of destination sheet used to set shtDest
    
    Dim frstRow As Integer: frstRow = 7
    Dim frstCol As Integer: frstCol = 22
    
    Dim lastRow As Integer: lastRow = shtSrc.Cells(shtSrc.Rows.Count, frstCol).End(xlUp).Row
    Dim lastCol As Integer: lastCol = 26
    
    Dim pstRow As Integer: pstRow = 308
    Dim pstCol As Integer: pstCol = 2
    
    'several MsgBox types to ensure user is doing exactly what they mean
    Dim continue
    continue = MsgBox("You are about to copy data into the following file:" & vbNewLine _
        & "     File: " & wrkSrc.Path & wrkPath & vbNewLine _
        & "     Sheet: " & shtDestNm & vbNewLine & vbNewLine _
        & "Are you sure you want to perform this action?", _
        4)
    If continue = vbYes Then
    
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        
        Set wrkDest = Workbooks.Open(ThisWorkbook.Path & wrkPath)
        Set shtDest = wrkDest.Sheets(shtDestNm)

        With shtSrc
            .Range(.Cells(frstRow, frstCol), .Cells(lastRow, lastCol)).Copy
        End With
        
        With shtDest
            .Range(.Cells(pstRow, pstCol), .Cells(pstRow + (lastRow - frstRow), pstCol + (lastCol - frstCol))).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        End With
            
        Application.CutCopyMode = False
        
        wrkDest.Activate
        
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        
        MsgBox ("Data has been copied into the destination sheet." & vbNewLine _
            & "Please review and save the file.")
    End If

End Sub

