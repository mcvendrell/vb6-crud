Attribute VB_Name = "HelpersMod"
Option Explicit

'-------------------------   FUNCIONES GENERALES DE AYUDA   -----------------------

'Esta funci�n devuelve el valor "Null" si la variable
'que se le pasa como par�metro esta vac�a. En caso contrario devuelve el
'valor de la variable. El valor devuelto ir� envuelto en comillas simples
'o no en el caso que sea de tipo num�rico (2� par�metro)
'Se hace la comprobaci�n con = "" porque si se le pasa Null devuelve cadena
'vac�a, al estar definido el par�metro como String
'Se declara como variant porque si se pasa un Rec Nulo, da error con string
Public Function Valor(strValor As Variant, Optional blnEsNumero As Boolean) As String
  If strValor = "" Or IsNull(strValor) Then
    Valor = "Null"
  Else
    Valor = IIf(blnEsNumero, strValor, "'" & strValor & "'")
  End If
End Function

'Parchea el texto pasado sustituyendo las ' por '' para grabarlos en SQL
Public Function TxtASql(strValor As String) As String
  TxtASql = Replace(strValor, "'", "''")
End Function

'-------------------------   CONTROL de ERRORES   -----------------------

'Proceso que muestra el error producido
'strMsg puede llevar un msg diferente del habitual, pero para que se muestre debe
'ir combinado con la opci�n blnMostrarSoloDescripcion, que a�adir� al final el error
Public Sub VerError(Optional blnMostrarSoloDescripcion As Boolean, Optional strMsg As String)
  strMsg = strMsg & IIf(blnMostrarSoloDescripcion, "", Err.Description)

  If blnMostrarSoloDescripcion Then
    MsgBox strMsg
  Else
    MsgBox "Se produjo un error: " & strMsg, vbInformation, "Error con la Base de Datos"
  End If
End Sub


'Funci�n que comprueba si ha habido alg�n error con la BD
'blnSinMensaje indica opcionalmente que s�lo compruebe el error sin mostrar ning�n msg indicativo
Public Function CompruebaError(Optional blnSinMensaje As Boolean) As Boolean
  CompruebaError = False

  'Se comprueba si ha habido alg�n error
  If Err.Number <> 0 Then
    If Not blnSinMensaje Then
      MsgBox "Se produjo un error: " & Err.Description, vbInformation, "Error con la Base de Datos"
    End If

    CompruebaError = True
    Exit Function
  End If
End Function
