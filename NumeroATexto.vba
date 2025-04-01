'=======================================================================
' Function: NumeroATexto
' Version: 1.5.7
' Author: Juan Pablo Garcia Murillo
' Date: 04/01/2025
' Description: 
'   This function converts a numeric value into a string representation in Spanish words. 
'   The function handles both the integer and decimal parts of the number, 
'   and formats it as a string in the format "X PESOS YY/100", where X is the integer 
'   part in words, and YY is the decimal part in two digits.
'   If the number exceeds 100,000 or is not numeric, an error message is returned.
'   The function also handles and removes commas and spaces from the input value.
'=======================================================================

Function NumeroATexto(ByVal MyNumber As Variant) As String
    ' Verifica que el valor sea numérico y lo limpia de comas y espacios.
    If Not IsNumeric(MyNumber) Then
        NumeroATexto = "ERROR: Valor no numérico"
        Exit Function
    End If
    MyNumber = Replace(CStr(MyNumber), ",", "") ' Remove commas
    MyNumber = Replace(MyNumber, " ", "") ' Remove spaces
    MyNumber = Val(MyNumber) ' Convert to numeric value
    
    ' Limitar a 100,000
    If MyNumber > 100000 Then
        NumeroATexto = "ERROR: El número excede el límite permitido (100,000)"
        Exit Function
    End If

    Dim entero As Long, decimales As Long
    entero = Int(MyNumber) ' Integer part
    decimales = Round((MyNumber - entero) * 100, 0) ' Decimal part (rounded to 2 decimal places)
    
    Dim resultado As String
    resultado = ""
    
    ' Convertir la parte entera (integer part)
    If entero = 0 Then
        resultado = "CERO"
    Else
        If entero >= 1000 Then
            Dim miles As Long
            miles = Int(entero / 1000)  ' This part could be 1 (for 1000-1999) or greater.
            If miles = 1 Then
                resultado = "MIL"
            Else
                resultado = ConvertirMenor1000(miles) & " MIL"
            End If
            entero = entero Mod 1000 ' Process the remainder (less than 1000)
            If entero > 0 Then
                resultado = resultado & " " & ConvertirMenor1000(entero)
            End If
        Else
            resultado = ConvertirMenor1000(entero) ' Convert integer part less than 1000
        End If
    End If
    
    ' Add the decimal part in the format "/100"
    resultado = resultado & " PESOS " & Format(decimales, "00") & "/100"
    
    ' Return the final result
    NumeroATexto = Application.Trim(resultado)
End Function
