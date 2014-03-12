Attribute VB_Name = "Refresh_All"
Sub RefreshAll()
    
    Dim start_time, end_time
    start_time = Now()
    Application.ScreenUpdating = False
    
    ThisWorkbook.RefreshAll
    
    Application.ScreenUpdating = True
    end_time = Now()
    
    MsgBox ("Refreshed all sheets in this workbook" & vbNewLine & "Time elapsed: " & DateDiff("s", start_time, end_time) & " seconds")
    
End Sub
