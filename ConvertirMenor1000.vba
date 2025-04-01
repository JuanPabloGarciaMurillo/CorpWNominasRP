'=======================================================================
' Function: ConvertirMenor1000
' Version: 1.5.7
' Author: Juan Pablo Garcia Murillo
' Date: 04/01/2025
' Description: 
'   This function converts a number less than 1000 into its Spanish word representation. 
'   It handles special cases (numbers from 0 to 19), numbers from 20 to 29 (e.g., "VEINTIUNO"), tens and units (e.g., "TREINTA Y CINCO"), and hundreds (e.g., "CIENTO VEINTE"). 
'   This function is used by `NumeroATexto` for converting integer values in numbers to words.
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
        result = especiales(n) ' Special cases (0 to 19)
        ConvertirMenor1000 = Application.Trim(result)
        Exit Function
    End If
    
    ' Manejar números entre 20 y 29: VEINTIUNO, VEINTIDOS, etc.
    If n < 30 Then
        If n = 20 Then
            result = "VEINTE" ' Exactly 20
        Else
            result = "VEINTI" & unidades(n - 20) ' Numbers between 21 and 29
        End If
        ConvertirMenor1000 = Application.Trim(result)
        Exit Function
    End If
    
    ' Para números entre 30 y 99 (30 to 99)
    If n < 100 Then
        Dim d As Long, u As Long
        d = Int(n / 10) ' Tens place
        u = n Mod 10 ' Units place
        If u = 0 Then
            result = decenas(d) ' Exact tens
        Else
            result = decenas(d) & " Y " & unidades(u) ' Tens with units
        End If
        ConvertirMenor1000 = Application.Trim(result)
        Exit Function
    End If
    
    ' Para números entre 100 y 999 (100 to 999)
    If n < 1000 Then
        Dim c As Long, resto As Long
        c = Int(n / 100) ' Hundreds place
        resto = n Mod 100 ' Remaining after hundreds
        If n = 100 Then
            result = "CIEN" ' Exactly 100
        Else
            If c = 1 Then
                result = "CIENTO" ' Special case for 100 to 199
            Else
                result = centenas(c) ' Hundreds
            End If
            If resto > 0 Then
                result = result & " " & ConvertirMenor1000(resto) ' Add the remainder
            End If
        End If
        ConvertirMenor1000 = Application.Trim(result)
        Exit Function
    End If
End Function
