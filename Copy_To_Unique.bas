Attribute VB_Name = "Copy_To_Unique"
Option Explicit

Sub cpToUnique()
    'This will copy an entire table (selected by contiguous rows in line 23) to a new sheet
    
    Dim start_time, end_time
    start_time = Now()
    
    Dim Source As Worksheet: Set Source = ThisWorkbook.Worksheets("cptest_src") 'source sheet
    Dim Destination As Worksheet: Set Destination = ThisWorkbook.Worksheets("cptest_dest") 'destination sheet
    
    Dim Records As Dictionary: Set Records = New Dictionary 'Dictionary will hold unique hashes as keys and the value will be the inde of the row in which that key was added
    
    Dim key As String 'must add because calling the array in a .Item call is not allowed
    Dim valToAdd As Variant 'must add because calling the array in a .Item call is not allowed
    Dim cpColBeg As Integer: cpColBeg = 7 'index of first column to add
    Dim cpColEnd As Integer: cpColEnd = 10 'index of final column to add
    Dim y As Integer 'y is used to iterate through columns when adding data to results
        
    Dim Data As Variant 'will hold source data as array
    Dim Index As Integer 'will refer to row in source data
    Dim Row As Integer: Row = 1 'indicates next row to copy
    
    Destination.Cells.Clear
    
    Data = Source.Range("A1").CurrentRegion.value
    
    For Index = LBound(Data, 1) To UBound(Data, 1) 'loops through all row indexes
        key = Data(Index, 1) 'value to check in dictionary
        
        If Records.Exists(key) Then
            For y = cpColBeg To cpColEnd 'iterates through columns to be added as specified and adds those values to existing value
                valToAdd = Data(Index, y)
                Destination.Cells(Records.Item(key), y).value = Destination.Cells(Records.Item(key), y).value + valToAdd
            Next y
        Else
            Records.Add key, Row 'adds hash and row index to dictionary
            
            For y = 1 To cpColEnd 'iterates through columns and adds to results in new line
                Destination.Cells(Row, y).value = Data(Index, y)
            Next y
            
            Row = Row + 1
        End If
    Next Index
    
    Set Records = Nothing
    
    end_time = Now()
    
    MsgBox ("Successfully unique-ified. Operation resulted in " & (Row - 1) & " rows." & vbNewLine & "Time elapsed: " & DateDiff("s", start_time, end_time) & " seconds")

End Sub
