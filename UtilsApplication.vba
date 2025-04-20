' Script: UtilsApplication
' Version: 0.9.3
' Author: Juan Pablo Garcia Murillo
' Date: 04/18/2025
' Description:
'   This module contains utility functions for managing application settings in Excel VBA.
' Functions included in this module:
'   - RestoreApplicationSettings
'   - OptimizeApplicationSettings

' Subroutine: RestoreApplicationSettings
' Description:
'   This subroutine restores the application settings to their default values.
' Notes:
'   - The subroutine sets ScreenUpdating to True, Calculation to Automatic, and clears CutCopyMode.
'   - It is used to restore the application settings after performing operations that may change them.

Public Sub RestoreApplicationSettings()
    Application.ScreenUpdating = TRUE
    Application.Calculation = xlCalculationAutomatic
    Application.CutCopyMode = FALSE
End Sub

' Subroutine: OptimizeApplicationSettings
' Description:
'   This subroutine optimizes the application settings for performance.
'   It turns off screen updating and sets calculation to manual.
'   This is useful for speeding up operations that involve multiple calculations or screen updates.
' Notes:
'   - The subroutine sets ScreenUpdating to False and Calculation to Manual.
'   - It is recommended to call RestoreApplicationSettings after performing operations that require optimization.

Public Sub OptimizeApplicationSettings()
    ' Turn off screen updating and set calculation to manual for performance
    Application.ScreenUpdating = FALSE
    Application.Calculation = xlCalculationManual
End Sub