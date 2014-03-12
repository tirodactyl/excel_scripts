Attribute VB_Name = "StringSplit_Function"
Function STR_SPLIT(str, sep, n) As String
    Dim V() As String
    V = Split(str, sep)
    STR_SPLIT = V(n - 1)
End Function
