Option Explicit

Function ConvertToTaka(ByVal value As Double) As String
    Const Thousand = 1000@
    Const Ten = 100@
    Const Lakh = Thousand * Ten
    Const Crore = Lakh * Ten
    Const Arab = Crore * Ten
    Const Kharab = Arab * Ten
    Const Neel = Kharab * Ten

    If value = 0@ Then
        ConvertToTaka = "Zero Taka Only"
        Exit Function
    End If

    Dim buf As String
    If value < 0@ Then
        buf = "Negative "
    Else
        buf = ""
    End If

    Dim frac As Double
    frac = Abs(value - Fix(value))
    If value < 0@ Or frac <> 0@ Then
        value = Abs(Fix(value))
    End If

    Dim atLeastOne As Integer
    atLeastOne = value >= 1@

    If value >= Neel Then
        buf = buf & SpellNumberDigitGroup(Int(value / Neel)) & " Neel"
        value = value - Int(value / Neel) * Neel
        If value >= 1@ Then buf = buf & " "
    End If

    If value >= Kharab Then
        buf = buf & SpellNumberDigitGroup(Int(value / Kharab)) & " Kharab"
        value = value - Int(value / Kharab) * Kharab
        If value >= 1@ Then buf = buf & " "
    End If

    If value >= Arab Then
        buf = buf & SpellNumberDigitGroup(Int(value / Arab)) & " Arab"
        value = value - Int(value / Arab) * Arab
        If value >= 1@ Then buf = buf & " "
    End If

    If value >= Crore Then
        buf = buf & SpellNumberDigitGroup(value \ Crore) & " Crore"
        value = value Mod Crore
        If value >= 1@ Then buf = buf & " "
    End If

    If value >= Lakh Then
        buf = buf & SpellNumberDigitGroup(value \ Lakh) & " Lac"
        value = value Mod Lakh
        If value >= 1@ Then buf = buf & " "
    End If

    If value >= Thousand Then
        buf = buf & SpellNumberDigitGroup(value \ Thousand) & " Thousand"
        value = value Mod Thousand
        If value >= 1@ Then buf = buf & " "
    End If

    If value >= 1@ Then
        buf = buf & SpellNumberDigitGroup(value) & " Taka"
    End If

    ' Add decimal places (Paisa)
    Dim paisa As Integer
    paisa = Int(frac * 100)
    If paisa > 0 Then
        buf = buf & " & " & SpellNumberDigitGroup(paisa) & " Paisa"
    End If

    buf = buf & " Only"

    ConvertToTaka = buf
End Function

Private Function SpellNumberDigitGroup(ByVal N As Integer) As String

   Const Hundred = " Hundred"
   Const One = "One"
   Const Two = "Two"
   Const Three = "Three"
   Const Four = "Four"
   Const Five = "Five"
   Const Six = "Six"
   Const Seven = "Seven"
   Const Eight = "Eight"
   Const Nine = "Nine"
   Dim buf As String: buf = ""
   Dim Flag As Integer: Flag = False

   Select Case (N \ 100)
      Case 0: buf = "": Flag = False
      Case 1: buf = One & Hundred: Flag = True
      Case 2: buf = Two & Hundred: Flag = True
      Case 3: buf = Three & Hundred: Flag = True
      Case 4: buf = Four & Hundred: Flag = True
      Case 5: buf = Five & Hundred: Flag = True
      Case 6: buf = Six & Hundred: Flag = True
      Case 7: buf = Seven & Hundred: Flag = True
      Case 8: buf = Eight & Hundred: Flag = True
      Case 9: buf = Nine & Hundred: Flag = True
   End Select

   If (Flag <> False) Then N = N Mod 100
   If (N > 0) Then
      If (Flag <> False) Then buf = buf & " "
   Else
      SpellNumberDigitGroup = buf
      Exit Function
   End If

   Select Case (N \ 10)
      Case 0, 1: Flag = False
      Case 2: buf = buf & "Twenty": Flag = True
      Case 3: buf = buf & "Thirty": Flag = True
      Case 4: buf = buf & "Forty": Flag = True
      Case 5: buf = buf & "Fifty": Flag = True
      Case 6: buf = buf & "Sixty": Flag = True
      Case 7: buf = buf & "Seventy": Flag = True
      Case 8: buf = buf & "Eighty": Flag = True
      Case 9: buf = buf & "Ninety": Flag = True
   End Select

   If (Flag <> False) Then N = N Mod 10
   If (N > 0) Then
      If (Flag <> False) Then buf = buf & "-"
   Else
      SpellNumberDigitGroup = buf
      Exit Function
   End If

   Select Case (N)
      Case 0:
      Case 1: buf = buf & One
      Case 2: buf = buf & Two
      Case 3: buf = buf & Three
      Case 4: buf = buf & Four
      Case 5: buf = buf & Five
      Case 6: buf = buf & Six
      Case 7: buf = buf & Seven
      Case 8: buf = buf & Eight
      Case 9: buf = buf & Nine
      Case 10: buf = buf & "Ten"
      Case 11: buf = buf & "Eleven"
      Case 12: buf = buf & "Twelve"
      Case 13: buf = buf & "Thirteen"
      Case 14: buf = buf & "Fourteen"
      Case 15: buf = buf & "Fifteen"
      Case 16: buf = buf & "Sixteen"
      Case 17: buf = buf & "Seventeen"
      Case 18: buf = buf & "Eighteen"
      Case 19: buf = buf & "Nineteen"
   End Select

   SpellNumberDigitGroup = buf

End Function
