'=======================================================================
' Script: ConvertirMenor1000
' Version: 1.5.6
' Author: Juan Pablo Garcia Murillo
' Date: 03/30/2025
' Description:
'   This helper function converts numbers less than 1000 into their Spanish
'   text equivalents. It handles special cases for numbers under 20, numbers
'   between 20 and 29, and numbers between 30 and 999, utilizing predefined
'   arrays for units, tens, and hundreds. It also ensures the proper formatting
'   of the number in text (e.g., "23" becomes "VEINTITRÉS").
'
'   Key functionalities:
'   - Converts numbers less than 1000 into Spanish text (e.g., "123" to "CIENTO VEINTITRÉS")
'   - Handles special cases for numbers below 20 and between 20 and 29
'   - Handles numbers between 30 and 999, considering units, tens, and hundreds
'   - Returns a properly formatted string for numbers in the range of 0 to 999
'   - Uses arrays for units, tens, and hundreds to efficiently convert numbers
'
'   This function is used as a part of the `NumeroATexto` function to convert
'   the integer portion of the number into its Spanish text representation.
'=======================================================================

Function ConvertirMenor1000(n As Long) As String
    ' Arreglos para números especiales y componentes
    Dim especiales As Variant
    especiales = Array("", "UNO", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE", _
                        "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE")
    
    Dim unidades As Variant, decenas As Variant, centenas As Variant
    unidades = Array("", "UNO", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE")
    decenas = Array("", "DIEZ", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA")
    centenas = Array("", "CIEN", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS")
    
    Dim result As String
    result = ""
    
    ' Si el número es menor a 20, usar el arreglo de especiales
    If n < 20 Then
        result = especiales(n)
        ConvertirMenor1000 = Application.Trim(result)
        Exit Function
    End If
    
    ' Manejar números entre 20 y 29: VEINTIUNO, VEINTIDOS, etc.
    If n < 30 Then
        If n = 20 Then
            result = "VEINTE"
        Else
            result = "VEINTI" & unidades(n - 20)
        End If
        ConvertirMenor1000 = Application.Trim(result)
        Exit Function
    End If
    
    ' Para números entre 30 y 99
    If n < 100 Then
        Dim d As Long, u As Long
        d = Int(n / 10)
        u = n Mod 10
        If u = 0 Then
            result = decenas(d)
        Else
            result = decenas(d) & " Y " & unidades(u)
        End If
        ConvertirMenor1000 = Application.Trim(result)
        Exit Function
    End If
    
    ' Para números entre 100 y 999
    If n < 1000 Then
        Dim c As Long, resto As Long
        c = Int(n / 100)
        resto = n Mod 100
        If n = 100 Then
            result = "CIEN"
        Else
            If c = 1 Then
                result = "CIENTO"
            Else
                result = centenas(c)
            End If
            If resto > 0 Then
                result = result & " " & ConvertirMenor1000(resto)
            End If
        End If
        ConvertirMenor1000 = Application.Trim(result)
        Exit Function
    End If
End Function

