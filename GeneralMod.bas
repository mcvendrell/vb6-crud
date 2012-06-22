Attribute VB_Name = "GeneralMod"
Option Explicit

'---------------------------------     VARIABLES     ------------------------------

'Conexion de la aplicación
'Conection to app
Public Conexion As ADODB.Connection

'Guarda la ruta desde donde se ejecuta el programa
'Path to the executing folder
Public GstrPath As String

'Función que trata de establecer una conexion con la base de datos
'Try to connect to the database
Public Function ConectaBD() As Boolean
  Dim strConexion As String
  
  Set Conexion = New ADODB.Connection
  
  On Local Error Resume Next
  
  strConexion = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & GstrPath & App.EXEName & ".mdb;UID=Administrador;PWD="
  Conexion.Open strConexion
  
  If CompruebaError() Then
    ConectaBD = False
    Set Conexion = Nothing
    
    MsgBox "Error intentando conectar con la BD. El error fue: " & Err.Description, vbCritical
  Else
    'Usar cliente sino no funciona un "Rec.AbsolutePosition"
    'We must use adUseClient or Rec.AbsolutePosition will not work
    Conexion.CursorLocation = adUseClient
    ConectaBD = True
  End If
End Function

'Salida del programa
'Exit from app
Public Sub SubSalir()
  On Local Error Resume Next
  Conexion.Close
  Set Conexion = Nothing
  End
End Sub

'Proceso Principal
Sub Main()
  'Obtener la ruta  de trabajo
  'Get path
  GstrPath = App.Path
  'Si se ejecuta en local no devuelve la última barra, pero en remoto si, comprobarlo
  'Ensure that path ends in \
  If Right(GstrPath, 1) <> "\" Then GstrPath = GstrPath & "\"
    
  If ConectaBD Then
    ClientesFrm.Show vbModal
  Else
    End
  End If
End Sub
