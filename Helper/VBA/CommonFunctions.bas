Attribute VB_Name = "CommonFunctions"
Const DimensionFormat As String = "0.000000"
Public Function Dim2Str(x As Double) As String
    Dim2Str = Format(x, DimensionFormat)
End Function

Public Function Form(ParamArray args()) As String
    Dim tmp As Variant
    Dim str As String
    
    For Each tmp In args
        str = str & CStr(tmp) & " "
    Next tmp
    Form = "(" & str & ")"
End Function

Public Function WFPi() As Double
    Pi = WorksheetFunction.Pi()
End Function

Public Function EscapeString(str As String) As String
    Dim tmp As String
    tmp = Replace(str, "\", "\\")
    tmp = Replace(tmp, """", "\""")
    EscapeString = tmp
End Function

