Attribute VB_Name = "GetRandom"
' ---------------------------------------------------------------------------
' Name: get_random_integer
' Description: You can get a random integer value.
' ---------------------------------------------------------------------------
Public Function get_random_integer( _
    ByVal lower As Integer, _
    ByVal upper As Integer _
) As Integer

    If lower > upper Then
        Exit Function
    End If

    get_random_integer = Int((upper - lower + 1) * Rnd + lower)

End Function

' ---------------------------------------------------------------------------
' Name: get_random_float
' Description: You can get a random float value.
' ---------------------------------------------------------------------------
Public Function get_random_float( _
    ByVal lower As Double, _
    ByVal upper As Double _
) As Double

    If lower > upper Then
        Exit Function
    End If

    get_random_float = (upper - lower) * Rnd + lower

End Function

' ---------------------------------------------------------------------------
' Name: get_random_char
' Description: You can get a random character.
' ---------------------------------------------------------------------------
Public Function get_random_char() As String

    Dim dest As Integer

    dest = get_random_integer(0, 61)

    If dest < 10 Then
        dest = dest + 48
    ElseIf dest < 36 Then
        dest = dest + 55
    Else
        dest = dest + 61
    End If

    get_random_char = Chr(dest)

End Function

' ---------------------------------------------------------------------------
' Name: get_random_string
' Description: You can get a random string.
' ---------------------------------------------------------------------------
Public Function get_random_string(ByVal length As Integer) As String

    Dim cat As String
    Dim i As Integer

    cat = get_random_char()
    For i = 2 To length
        cat = cat + get_random_char()
    Next i

    get_random_string = cat

End Function
