Attribute VB_Name = "AssetCodes"
Public Function isin_checkDigit(isin As String) As Integer

    Dim even As Integer
    Dim odd As Integer
    Dim length As Integer

    even = 0
    odd = 0
    length = Len(isin)
    
    If length <> 11 Then
        MsgBox "Wrong length"
        Exit Function
    End If
    
    For i = 1 To length
        digit = Mid(isin, i, 1)
        If IsNumeric(digit) Then
            temp = temp & digit
        Else
         temp = temp & (Asc(digit) - Asc("A") + 10)
        End If
    Next i
    
    For i = Len(temp) To 1 Step -1
        digit = CInt(Mid(temp, i, 1))
        
        If Len(temp) Mod 2 = 0 Then
            If i Mod 2 <> 0 Then
                odd = odd + digit
            Else
                digit = digit * 2
                If digit > 9 Then
                    digit = digit - 9
                End If
                even = even + digit
            End If
        Else
            If i Mod 2 = 0 Then
                odd = odd + digit
            Else
                digit = digit * 2
                If digit > 9 Then
                    digit = digit - 9
                End If
                even = even + digit
            End If
        End If
    Next i
    
    isin_checkDigit = (10 - ((odd + even) Mod 10)) Mod 10

End Function


Public Function wknToIsin(wkn As String) As String

    Dim isin As String
    
    isin = "DE000" & wkn
    wknToIsin = isin & isin_checkDigit(isin)
    
End Function

Public Function isinToWkn(isin As String) As String

    Dim wkn As String
    
    isinToWkn = wkn & Mid(isin, 6, 6)
    
End Function
