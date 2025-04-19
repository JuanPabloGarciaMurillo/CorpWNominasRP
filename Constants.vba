'=========================================================
' Module: Constants
' Version: 0.9.1
' Author: Juan Pablo Garcia Murillo
' Date: 04/18/2025
' Description:
'   This module contains all the public constants used across the project.
'   These constants are shared across multiple modules to ensure consistency, improve maintainability, and avoid hardcoding values throughout the codebase.
' Usage:
'   - Declare all constants as `Public Const` to make them accessible across the entire project.
'   - Group related constants logically (e.g., sheet names, table names, column names, error messages).
'=========================================================

'Sheet Constants
Public Const PLANTEL_SHEET As String = "Planteles"
Public Const PREMIOS_SHEET As String = "Premios"
Public Const CURSOS_SHEET As String = "Cursos"
Public Const TABULADORES_SHEET As String = "Tabuladores"
Public Const COLABORADORES_SHEET As String = "Colaboradores"
Public Const RESULTADOS_SHEET As String = "Resultados"
Public Const DASHBOARD_SHEET As String = "Dashboard"
Public Const COORDINADORES_SHEET As String = "Coordinadores"
Public Const PROMOTORES_SHEET As String = "Promotores"
Public Const SKIP_SHEETS As String = PLANTEL_SHEET & "," & PREMIOS_SHEET & "," & CURSOS_SHEET & "," & TABULADORES_SHEET & "," & COLABORADORES_SHEET & "," & RESULTADOS_SHEET & "," & DASHBOARD_SHEET & "," & COORDINADORES_SHEET & "," & PROMOTORES_SHEET

' Table Constants
Public Const COORDINADOR_TABLE As String = "Tabla_Coordinador"
Public Const COORDINADORES_TABLE As String = "Tabla_Coordinadores"
Public Const ACTIVE_TABLE As String = "Tabla_Coordinadores_Gerencia_Activa"
Public Const CURSOS_TABLE As String = "Tabla_Cursos"
Public Const DESCUENTOS_TABLE As String = "Tabla_Descuentos"
Public Const GERENTE_TABLE As String = "Tabla_Gerente"
Public Const GERENTES_TABLE As String = "Tabla_Gerentes"
Public Const PLANTELES_TABLE As String = "Tabla_Planteles"
Public Const PREMIOS_COORDINADOR_TABLE As String = "Tabla_Premios_Coordinador"
Public Const PREMIOS_GERENTE_TABLE As String = "Tabla_Premios_Gerente"
Public Const PREMIOS_PROMOTOR_TABLE As String = "Tabla_Premios_Promotor"
Public Const PROMOTOR_TABLE As String = "Tabla_Promotor"
Public Const PROMOTORES_TABLE As String = "Tabla_Promotores"
Public Const RAZONES_SOCIALES_TABLE As String = "Tabla_Razones_Sociales"
Public Const SUELDO_BASE_POR_PUESTO_TABLE As String = "Tabla_Sueldo_Base_Por_Puesto"
Public Const SUELDOS_BASE_TABLE As String = "Tabla_Sueldos_Base"
Public Const TABULADOR_COORDINADOR_TABLE As String = "Tabla_Tabulador_Coordinador"
Public Const TABULADOR_GERENTE_TABLE As String = "Tabla_Tabulador_Gerente"
Public Const TABULADOR_PREMIO_PROMOTOR_TOTAL_GERENCIA_TABLE As String = "Tabla_Tabulador_Premio_Promotor_Total_Gerencia"
Public Const TABULADOR_PROMOTOR_TABLE As String = "Tabla_Tabulador_Promotor"

' Column Constants
Public Const GERENCIA_COLUMN As String = "GERENCIA"
Public Const COORDINADOR_COLUMN As String = "COORDINADOR"
Public Const PROMOTOR_COLUMN As String = "PROMOTOR"
Public Const NOMBRE_COLUMN As String = "NOMBRE"
Public Const COORDINACION_COLUMN As String = "COORDINACION"
Public Const COLABORADOR_COLUMN As String = "COLABORADOR"
Public Const ALIAS_COLUMN As String = "ALIAS"
Public Const SHEET_NAMES_COLUMN As String = "P"
Public Const START_COLUMN As Long = 2
Public Const COLUMN_A As String = "A"
Public Const COLUMN_D As String = "D"
Public Const COLUMN_E As String = "E"

'Text Constants
Public Const PAGO_NETO_TEXT As String = "PAGO NETO"
Public Const ZERO_TEXT As String = "CERO"
Public Const PESOS_TEXT As String = "PESOS"

' Properties Constants
Public Const HEADERS As String = "PROMOTOR,CREDENCIAL,NOMBRE DEL ALUMNO,PLANTEL,CURSO,GRUPO,FECHA,TS PLANTEL,TS CREDENCIAL"
Public Const COLUMN_INDICES As String = "1,2,3,5,6,7,9,10,11"
Public Const TAB_SUFFIX As String = "(C)"
Public Const MAX_LIMIT As Long = 100000
Public Const TARGET_CELL As String = "J4"
Public Const MANAGER_IDENTIFIER As String = "GERENTE"

' Error Message Constants
Public Const ERROR_SHEET_NOT_FOUND As String = "Hoja '"
Public Const ERROR_GENERIC As String = "Error. Porfavor contacta a tu administrador: "
Public Const ERROR_INVALID_SHEET As String = "La hoja destino no es válida."
Public Const ERROR_NO_VALID_MANAGER As String = "No se encontraron gerentes válidos. Saliendo de la macro."
Public Const ERROR_NO_VALID_COORDINATOR As String = "No se encontraron coordinadores válidos. Saliendo de la macro."
Public Const ERROR_EMPTY_MANAGER_CELL As String = "La celda 'Nombre_Gerente' está vacía. Por favor, ingrese un nombre de gerente válido."
Public Const ERROR_NO_COORDINATORS As String = "No se encontraron coordinadores para el gerente '"
Public Const ERROR_UPDATE_TABLE As String = "Error al actualizar la tabla Coordinadores_Gerencia_Activa: "
Public Const ERROR_UPDATE_DASHBOARD As String = "Error al actualizar la tabla Coordinadores_Gerencia_Activa o los gráficos del Dashboard: "
Public Const ERROR_KEY_EXISTS As String = "El valor ya existe: "
Public Const ERROR_KEY_NOT_FOUND As String = "Valor no encontrado: "
