Attribute VB_Name = "StringSplit_Function"
Function STR_SPLIT(ByVal Rng As Range, ByVal sDelimiter As String, ByVal iPos As Integer) As String
    Dim vSplit As Variant
    'Application.Volatile
    vSplit = Split(Rng.Value, sDelimiter)
    STR_SPLIT = vSplit(iPos - 1)
End Function

