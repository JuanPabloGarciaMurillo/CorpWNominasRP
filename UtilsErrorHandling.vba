'=========================================================
' Script: UtilsErrorHandling
' Version: 0.9.2
' Author: Juan Pablo Garcia Murillo
' Date: 04/18/2025
' Description:
'   This module contains utility functions for error handling in Excel VBA.
'   It includes functions for displaying error messages, logging errors, and handling specific error types.
' Functions included in this module:
'   - HandleError
'=========================================================

'=========================================================
' Subroutine: HandleError
' Description:
'   Displays an error message box with a generic error message and optionally logs the source of the error.
' Parameters:
'   - errorMessage (String): The error message to display.
'   - errorSource (String, Optional): The name of the file or subroutine that caused the error.
' Returns:
'   - None
' Notes:
'   - The function uses a message box to display the error message.
'   - The message box is displayed with a critical icon and an "Error" title.
'=========================================================
Public Sub HandleError(errorMessage As String, Optional errorSource As String = "")
    Dim fullMessage As String
    
    ' Construct the full error message
    If errorSource <> "" Then
        fullMessage = "Error in [" & errorSource & "]: " & errorMessage
    Else
        fullMessage = errorMessage
    End If
    
    ' Display the error message
    MsgBox ERROR_GENERIC & fullMessage, vbCritical, "Error"
End Sub