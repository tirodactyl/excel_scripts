Attribute VB_Name = "Consolidate_Data_P"
Option Explicit
 
Sub ConsDataP()
     
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
    Dim x As Integer 'used for iteration when adding hash
    Dim finalRow As Integer
    
    Application.ScreenUpdating = False
     
    dataInd = "P_" 'Characters placed at the beginning of a sheet name to indicate that this is a data sheet to be compiled
    Set wrkConsSheet = Sheets("Combined_Partner") 'Sheet (tab) name to consolidate data. Change to suit.
    wrkConsTblNm = "CombPartnerTable" 'Name of resulting table in sheet named above (allows for a pivot to reference the table instead of a specified range)
    srcClBeg = "E" 'Column to begin copying values in data sheets
    srcClEnd = "N" 'Column to end copying values in data sheets
    pstClBeg = "A" 'Column to begin pasting into consolidate sheet (used for searching and table creation)
    pstClEnd = "K" 'Final pasted column in consolidated sheet (used for table creation)
    
    wrkConsSheet.Cells.Clear 'Clear sheet for clean run each time
    
    lngMyCounter = 0 'Resets each time you run the macro
    
    For Each wrkMySheet In ThisWorkbook.Sheets 'Loops for all worksheets
        wrkName = wrkMySheet.Name
        If InStr(1, wrkName, dataInd) = 1 Then 'Brings in any sheet with a specific prefix indicating it is a data sheet (based on dataInd)
            lngLastRow = wrkMySheet.Range(srcClBeg & ":" & srcClEnd).Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
            If lngMyCounter = 0 Then
                wrkMySheet.Range(srcClBeg & "1:" & srcClEnd & lngLastRow).Copy 'Copy from sheet
                wrkConsSheet.Range(pstClBeg & "1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False 'Paste values w/format
            Else
                lngOutputRow = wrkConsSheet.Range(pstClBeg & ":" & pstClEnd).Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row + 1
                wrkMySheet.Range(srcClBeg & "2:" & srcClEnd & lngLastRow).Copy 'Copy from sheet
                wrkConsSheet.Range(pstClBeg & lngOutputRow).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False 'Paste values w/format
            End If
            lngMyCounter = lngMyCounter + 1 'counts number of sheets combined
        End If
    Next wrkMySheet
    
    finalRow = wrkConsSheet.Range(pstClBeg & ":" & pstClBeg).Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    'following block inserts a column that creates a hash from other columns in the sheet
    wrkConsSheet.Range("A1").EntireColumn.Insert
    wrkConsSheet.Range("A1").value = "HASH"
    
    For x = 2 To finalRow
        
        '***THIS BLOCK IS BECAUSE WE CANNOT DEAL WITH AN 'UNKNOWN' DEVICE - WE DEFAULT TO PHONE***
        If InStr(wrkConsSheet.Range("E" & x).value, "unknown") Then
            wrkConsSheet.Range("E" & x).value = "Phone"
        End If
        '***THIS BLOCK CREATES A HASH FOR MATCHING PURPOSES
        '*******THE IF STATEMENT IS BECAUSE WE CANNOT BREAK OUT MOBILE FACEBOOK PARTNERS YET IN KONTAGENT AND THEREFORE HAVE TO LUMP IT ALL TOGETHER - NOTE THAT IT RETAINS PARTNER NAME TO ENABLE BREAKOUT PRIOR TO MATCHING KT***
        If InStr(wrkConsSheet.Range("C" & x).value, "Facebook") Then
            If Not InStr(wrkConsSheet.Range("D" & x).value, "Canvas") Then
                wrkConsSheet.Range("A" & x).value = "Facebook" + wrkConsSheet.Range("D" & x).value + wrkConsSheet.Range("E" & x).value + wrkConsSheet.Range("F" & x).value + Format(wrkConsSheet.Range("G" & x).value, "YYYYMMDD")
            End If
        Else
            wrkConsSheet.Range("A" & x).value = wrkConsSheet.Range("B" & x).value + wrkConsSheet.Range("D" & x).value + wrkConsSheet.Range("E" & x).value + wrkConsSheet.Range("F" & x).value + Format(wrkConsSheet.Range("G" & x).value, "YYYYMMDD")
        End If
    Next x
    
    'END hash insert
    
    didWin = createTable(wrkConsSheet.Name, pstClBeg & "1:" & pstClEnd & finalRow, wrkConsTblNm)
    
    Dim b As Variant
    Dim c As Variant
    
    b = wrkConsSheet.Range(wrkConsTblNm & "[P_Date]").Select
    
    For Each c In b
        c = Format(Date, "m/d/yy")
    Next
    
    Application.ScreenUpdating = True
    end_time = Now()
    
    If didWin = True Then
        MsgBox ("Number of sheets consolidated: " & lngMyCounter & vbNewLine _
            & "Destination table: " & wrkConsTblNm & vbNewLine _
            & "     Table range: " & pstClBeg & "1:" & pstClEnd & finalRow & vbNewLine _
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
