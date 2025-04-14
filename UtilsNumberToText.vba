'==================================================
' Script: UtilsNumberToText
' Author: Juan Pablo Garcia Murillo
' Date: 04/06/2025
' Description:
'   This module contains utility functions for converting numbers to their Spanish word representation, specifically for financial amounts in pesos.
'   It includes functions for converting numbers to text, handling special cases, and validating input values. The module is designed to work with Excel VBA, making it easier to format and display numeric values in a human-readable format.
'
' Functions included in this module:
'   - NumeroATexto
'   - ConvertirMenor1000
'==================================================

'==================================================
' Function: NumeroATexto
' Description:
'   This function converts a numeric value into its Spanish word representation,
'   formatted       as an amount in pesos with two decimal places.
'   The function handles numbers up to 100,000 and includes the decimal part in "/100" format.
' Parameters:
'   - MyNumber (Variant): The number to convert to text.
' Returns:
'   - String: The Spanish text representation of the number, including "PESOS XX/100".
' Notes:
'   - Removes commas and spaces before processing.
'   - Returns an error message if the input is non-numeric or exceeds 100,000.
'   - Uses `ConvertirMenor1000` to process numbers less than 1000.
'==================================================

Public Function NumeroATexto(ByVal MyNumber As Variant) As String
    ' Verifica que el valor sea numérico y lo limpia de comas y espacios.
    If Not IsNumeric(MyNumber) Then
        NumeroATexto = "ERROR: Valor no numérico"
        Exit Function
    End If
    MyNumber = Replace(CStr(MyNumber), ",", "")        ' Remove commas
    MyNumber = Replace(MyNumber, " ", "")        ' Remove spaces
    MyNumber = Val(MyNumber)        ' Convert to numeric value
    
    ' Limitar a 100,000
    If MyNumber > 100000 Then
        NumeroATexto = "ERROR: El número excede el límite permitido (100,000)"
        Exit Function
    End If
    
    Dim entero      As Long, decimales As Long
    entero = Int(MyNumber)        ' Integer part
    decimales = Round((MyNumber - entero) * 100, 0)        ' Decimal part (rounded to 2 decimal places)
    
    Dim resultado   As String
    resultado = ""
    
    ' Convertir la parte entera (integer part)
    If entero = 0 Then
        resultado = "CERO"
    Else
        If entero >= 1000 Then
            Dim miles As Long
            miles = Int(entero / 1000)        ' This part could be 1 (for 1000-1999) or greater.
            If miles = 1 Then
                resultado = "MIL"
            Else
                resultado = ConvertirMenor1000(miles) & " MIL"
            End If
            entero = entero Mod 1000        ' Process the remainder (less than 1000)
            If entero > 0 Then
                resultado = resultado & " " & ConvertirMenor1000(entero)
            End If
        Else
            resultado = ConvertirMenor1000(entero)        ' Convert integer part less than 1000
        End If
    End If
    
    ' Add the decimal part in the format "/100"
    resultado = resultado & " PESOS " & Format(decimales, "00") & "/100"
    
    ' Return the final result
    NumeroATexto = Application.Trim(resultado)
End Function

'==================================================
' Function: ConvertirMenor1000
' Description:
'   This function converts a number less than 1000 into its Spanish word representation.
'   It handles special cases (numbers from 0 to 19), numbers from 20 to 29 (e.g., "VEINTIUNO"),
'   tens and units (e.g., "TREINTA Y CINCO"), and hundreds (e.g., "CIENTO VEINTE").
'   This function is used by `NumeroATexto` for converting integer values in numbers to words.
' Parameters:
'   - n (Long): The number to convert (must be between 0 and 999).
' Returns:
'   - String: The Spanish word representation of the number.
' Notes:
'   - Uses predefined arrays for special cases, units, tens, and hundreds.
'   - Handles unique Spanish numbering rules such as "VEINTIUNO" instead of "VEINTE Y UNO".
'   - "CIEN" is used for exactly 100, while "CIENTO" is used for numbers like 101-199.
'==================================================

Public Function ConvertirMenor1000(n As Long) As String
    ' Arreglos para números especiales y componentes
    Dim especiales  As Variant
    especiales = Array("", "UNO", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE", _
                 "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE")
    
    Dim unidades    As Variant, decenas As Variant, centenas As Variant
    unidades = Array("", "UNO", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE")
    decenas = Array("", "DIEZ", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA")
    centenas = Array("", "CIEN", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS")
    
    Dim result      As String
    result = ""
    
    ' Si el número es menor a 20, usar el arreglo de especiales
    If n < 20 Then
        result = especiales(n)        ' Special cases (0 to 19)
        ConvertirMenor1000 = Application.Trim(result)
        Exit Function
    End If
    
    ' Manejar números entre 20 y 29: VEINTIUNO, VEINTIDOS, etc.
    If n < 30 Then
        If n = 20 Then
            result = "VEINTE"        ' Exactly 20
        Else
            result = "VEINTI" & unidades(n - 20)        ' Numbers between 21 and 29
        End If
        ConvertirMenor1000 = Application.Trim(result)
        Exit Function
    End If
    
    ' Para números entre 30 y 99 (30 to 99)
    If n < 100 Then
        Dim d       As Long, u As Long
        d = Int(n / 10)        ' Tens place
        u = n Mod 10        ' Units place
        If u = 0 Then
            result = decenas(d)        ' Exact tens
        Else
            result = decenas(d) & " Y " & unidades(u)        ' Tens with units
        End If
        ConvertirMenor1000 = Application.Trim(result)
        Exit Function
    End If
    
    ' Para números entre 100 y 999 (100 to 999)
    If n < 1000 Then
        Dim c       As Long, resto As Long
        c = Int(n / 100)        ' Hundreds place
        resto = n Mod 100        ' Remaining after hundreds
        If n = 100 Then
            result = "CIEN"        ' Exactly 100
        Else
            If c = 1 Then
                result = "CIENTO"        ' Special case for 100 to 199
            Else
                result = centenas(c)        ' Hundreds
            End If
            If resto > 0 Then
                result = result & " " & ConvertirMenor1000(resto)        ' Add the remainder
            End If
        End If
        ConvertirMenor1000 = Application.Trim(result)
        Exit Function
    End If
End Function