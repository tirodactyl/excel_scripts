Attribute VB_Name = "Copy_From_UATracker"
Sub cpFrmUATracker() 'copies sheet from another workbook to this one

    Dim start_time, end_time
    
    Dim wrkDest As Workbook: Set wrkDest = ThisWorkbook
    Dim shtDest As Worksheet
    Dim shtDestNm As String
    Dim wrkSrc As Workbook 'source workbook opened to paste data into
    Dim wrkPath As String 'path of source workbook including file name - needed in order to open the file (MUST BE IN THE SAME PARENT FOLDER)
    Dim shtSrc As Worksheet 'source sheet to copy data into
    Dim shtSrcNm As String 'string representing the name of the source worksheet
    Dim shtCopy As Worksheet
    Dim destTblNm As String
    
    wrkPath = ":" & "UA_Tracker_v.93.xlsm" 'path within parent directory to the source workbook ( ":" instead of "/")
    shtSrcNm = "Matched_Data" 'name of source sheet used to set shtSrc
    shtDestNm = "Matched_Data" 'name of destination sheet
    destTblNm = "MatchedData"
    
    
    'MsgBox ensures user is doing exactly what they mean
    Dim continue
    continue = MsgBox("You are about to copy data from the following file:" & vbNewLine _
        & "     File: " & ThisWorkbook.Path & wrkPath & vbNewLine _
        & "     Sheet: " & shtSrcNm & vbNewLine & vbNewLine _
        & "This operation cannot be reversed." & vbNewLine _
        & "Are you sure you want to overwrite sheet |" & shtDestNm & "| at this time?" & vbNewLine & vbNewLine _
        & "**NOTE: You must select to enable macros when this script attempts to open the source file.", _
        4)

    If continue = vbYes Then
        start_time = Now()
        
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
    
        Set wrkSrc = Workbooks.Open(ThisWorkbook.Path & wrkPath)
        Set shtSrc = wrkSrc.Sheets(shtSrcNm)
        Set shtDest = wrkDest.Sheets(shtDestNm)
        shtDest.Cells.Clear
        
        shtSrc.Range("A1").CurrentRegion.Copy
        
        shtDest.Range("A1").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        Application.CutCopyMode = False
        
        wrkSrc.Close (False)
        
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        
        end_time = Now()
        
        MsgBox ("Sheet |" & shtDestNm & "| successfully overwritten!" & vbNewLine _
            & "     Please remember to save this file." & vbNewLine _
            & "Time elapsed: " & DateDiff("s", start_time, end_time) & " seconds")
    End If
    
    wrkDest.Sheets("AdX_Dump").Activate
     
End Sub
