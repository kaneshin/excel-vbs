Attribute VB_Name = "run_GetRandom"
Sub run_GetRandom()
    Range("B2") = GetRandom.get_random_integer(-5, 10)
    Range("B3") = GetRandom.get_random_double(-2.5, 7.5)
    Range("B4") = GetRandom.get_random_char()
    Range("B5") = GetRandom.get_random_string(8)
End Sub
