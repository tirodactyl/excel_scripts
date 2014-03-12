Attribute VB_Name = "Consolidate_Data"
Option Explicit
 
Sub ConsData()
     
     'This script started life as the first entry on: http://www.ozgrid.com/forum/showthread.php?t=167716
    
    Dim start_time, end_time
    start_time = Now()
    
    Dim wrkMySheet As Worksheet, _
    wrkConsSheet As Worksheet
    Dim lngLastRow As Long, _
    lngOutputRow As Long, _
    lngMyCounter As Long
    Dim wrkName As String, _
    wrkConsTblNm As String, _
    dataInd As String, _
    srcRange As String, _
    srcClBeg As String, _
    srcClEnd As String, _
    pstClBeg As String, _
    pstClEnd As String, _
    didWin As String

    
    Application.ScreenUpdating = False
     
    Set wrkConsSheet = Sheets("Combined Data Sheet") 'Sheet (tab) name to consolidate data. Change to suit.
    wrkConsTblNm = "CombDataTable" 'Name of resulting table in sheet named above (allows for a pivot to reference the table instead of a specified range)
    srcClBeg = "E" 'Column to begin copying values in data sheets
    srcClEnd = "M" 'Column to end copying values in data sheets
    pstClBeg = "A" 'Column to begin pasting into consolidate sheet (used for searching)
    pstClEnd = "I" 'Final pasted column in consolidated sheet (used for searching)
    
    wrkConsSheet.Cells.Clear 'Clear sheet for clean run each time
    
    lngMyCounter = 0 'Resets each time you run the macro
    
    dataInd = "P_" 'Characters placed at the beginning of a sheet name to indicate that this is a data sheet to be compiled
    
    For Each wrkMySheet In ThisWorkbook.Sheets 'Loops for all worksheets
        wrkName = wrkMySheet.Name
        If InStr(1, wrkName, dataInd) = 1 Then 'Brings in any sheet with a specific prefix indicating it is a data sheet (based on dataInd)
            lngLastRow = wrkMySheet.Range("E:M").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
            If lngMyCounter = 0 Then
                wrkMySheet.Range(srcClBeg & "1:" & srcClEnd & lngLastRow).Copy 'Copy from sheet
                wrkConsSheet.Range(pstClBeg & "1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False 'Paste values w/format
            Else
                lngOutputRow = wrkConsSheet.Range(pstClBeg & ":" & pstClEnd).Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row + 1
                wrkMySheet.Range(srcClBeg & "2:" & srcClEnd & lngLastRow).Copy 'Copy from sheet
                wrkConsSheet.Range(pstClBeg & lngOutputRow).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False 'Paste values w/format
            End If
            lngMyCounter = lngMyCounter + 1
        End If
    Next wrkMySheet
    
    didWin = createTable(wrkConsSheet.Name, pstClBeg & "1:" & pstClEnd & wrkConsSheet.Range(pstClBeg & ":" & pstClEnd).Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row, wrkConsTblNm)
    
    Application.ScreenUpdating = True
    end_time = Now()
    
    If didWin = True Then
        MsgBox ("Number of sheets consolidated: " & lngMyCounter & vbNewLine _
            & "Destination table: " & wrkConsTblNm & vbNewLine _
            & "     Table range: " & pstClBeg & "1:" & pstClEnd & wrkConsSheet.Range(pstClBeg & ":" & pstClEnd).Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row & vbNewLine _
            & "Destination sheet: " & wrkConsSheet.Name & vbNewLine _
            & "Time elapsed: " & DateDiff("s", start_time, end_time) & " seconds")
    End If
     
End Sub

Function createTable(shtNm As String, tblRange As String, tblNm As String) As Boolean
    'When called includes the following params:
    'shtNm - name of the sheet in which table will be created
    'tblRange - range to be included in the table, as a string
    'tblNm - string to use as name for the table (alphanumeric only)
 
    'Create Table in Excel VBA
    Sheets(shtNm).ListObjects.Add(xlSrcRange, Sheets(shtNm).Range(tblRange), , xlYes).Name = tblNm
    createTable = True

End Function
