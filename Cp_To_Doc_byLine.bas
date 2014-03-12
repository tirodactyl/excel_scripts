Attribute VB_Name = "Cp_To_Doc"
Sub cpToOldDoc() 'copies data in a pivot table to a sheet in another workbook

    Dim wrkSrc As Workbook: Set wrkSrc = ThisWorkbook
    Dim shtSrc As Worksheet: Set shtSrc = ThisWorkbook.Sheets("Legacy Update")
    Dim wrkDest As Workbook 'destination workbook opened to paste data into
    Dim wrkPath As String 'path of destination workbook including file name - needed in order to open the file (MUST BE IN THE SAME PARENT FOLDER)
    Dim shtDest As Worksheet 'destination sheet to copy data into
    Dim shtDestNm As String 'string representing the name of the destination worksheet
    
    wrkPath = ":" & "UA_copy_test.xlsx" 'path within parent directory to the destination workbook ( ":" instead of "/")
    shtDestNm = "MADD" 'name of destination sheet used to set shtDest
    
    Dim pstCol As Integer: pstCol = 2
    Dim pstRow As Integer: pstRow = 7
    
    Dim x As Integer
    Dim y As Integer
    
    Dim tblToCopy As Variant 'this will be an array made from a range in the source sheet; we will iterate through it to copy values
    tblToCopy = shtSrc.Range("B7").CurrentRegion.value
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Set wrkDest = Workbooks.Open(ThisWorkbook.Path & wrkPath)
    Set shtDest = wrkDest.Sheets(shtDestNm)
    
    '
    Dim continue
    continue = MsgBox("You are about to copy data into the following file:" & vbNewLine _
        & "     File: " & wrkDest.Path & vbNewLine _
        & "     Sheet: " & shtDest.Name & vbNewLine & vbNewLine _
        & "Are you sure you want to perform this action?", _
        4)
    If continue = vbYes Then
        With shtDest
            For x = 3 To UBound(tblToCopy, 1) 'loops through all row indices (NOTE: starts at 3 due to pivot table using two rows for headers)
                    For y = 2 To UBound(tblToCopy, 2) 'loops through all column indices (NOTE: starts at 2 due to pivot table using one column for row label)
                        .Cells(pstRow + (x - 3), pstCol + (y - 2)).value = tblToCopy(x, y)
                    Next y
            Next x
            .Range(Cells(pstRow, pstCol), Cells(pstRow + (UBound(tblToCopy, 1) - 1), pstCol + (UBound(tblToCopy, 2) - 1))).Select
        End With
        
        wrkSrc.Windows(1).WindowState = xlMinimized
        shtDest.Activate
        wrkDest.Windows(1).WindowState = xlMaximized
        
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        
        Dim saveIt
        saveIt = MsgBox("Please check behind this box to ensure the data has been correctly copied into the destination sheet." & vbNewLine _
            & "Would you like to save the file?", _
            4)
        
        If saveIt = vbYes Then
            wrkDest.Close (True)
        Else
            wrkDest.Close (False)
        End If
        
        wrkSrc.Windows(1).WindowState = xlMaximized
    Else
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        wrkDest.Close (False)
    End If

End Sub
