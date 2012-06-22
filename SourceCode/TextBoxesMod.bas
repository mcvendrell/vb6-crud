Attribute VB_Name = "TextBoxesMod"
Option Explicit

'-------------------------   CONTROL de TEXTBOXES   -----------------------

'Comprueba que la tecla pulsada corresponde a un carácter para fecha válido
Public Function CompruebaFecha(intTecla As Integer) As Integer
  Select Case intTecla
    Case Asc("0") To Asc("9")
      'Se ha pulsado un nº
    Case Asc("/")
      'Se ha pulsado un separador
    Case Asc("."), Asc("-")
      'Se ha pulsado un separador, cambiar por :
      intTecla = Asc("/")
    Case 8
      'Se ha pulsado BCKSPC
    Case Else
      intTecla = 0
      Beep
  End Select
  
  CompruebaFecha = intTecla
End Function

'Comprueba que la tecla pulsada corresponde a un nº (para enteros porque
'no permite la pulsación de la coma
Public Function CompruebaEntero(intTecla As Integer) As Integer
  Select Case intTecla
    Case Asc("0") To Asc("9")
      'Se ha pulsado un nº
    Case 8
      'Se ha pulsado BCKSPC
    Case Else
      'Si no son nºs o BCKSPC, no se admite
      intTecla = 0
      Beep
  End Select
  
  CompruebaEntero = intTecla
End Function

'Comprueba que la tecla pulsada corresponde a un nº (para enteros porque
'no permite la pulsación de la coma
Public Function CompruebaEnteroConSigno(intTecla As Integer) As Integer
  Select Case intTecla
    Case Asc("0") To Asc("9")
      'Se ha pulsado un nº
    Case 8, 45
      'Se ha pulsado BCKSPC o -
    Case Else
      'Si no son nºs, el - o BCKSPC, no se admite
      intTecla = 0
      Beep
  End Select
  
  CompruebaEnteroConSigno = intTecla
End Function

'Comprueba que la tecla pulsada corresponde a un nº, decimal o no
Public Function CompruebaNumero(intTecla As Integer, strTexto As String) As Integer
  Select Case intTecla
    Case Asc("0") To Asc("9")
      'Se ha pulsado un nº
    
    Case Asc("."), Asc(",")
      Select Case intTecla
        'Case (Asc(".") And FormatoNumericoOk)
          'Se ha pulsado el punto y es el separador decimal correcto
        
        Case (Asc(".") And Not FormatoNumericoOk)
          'Se ha pulsado el punto y NO es el separador decimal correcto
          'Pasar a coma si no hay otra
          intTecla = Asc(",")
        
        Case (Asc(",") And FormatoNumericoOk)
          'Se ha pulsado la coma y NO es el separador decimal correcto
          'Pasar a punto si no hay otro
          intTecla = Asc(".")
      
      End Select
      
      If InStr(strTexto, Chr(intTecla)) > 0 Then
        'Ya existe un punto en la cadena
        intTecla = 0
        Beep
      End If
      
      
    Case 8
      'Se ha pulsado BCKSPC
    Case Else
      'Si no son nºs, el punto o BCKSPC, no se admite
      intTecla = 0
      Beep
  End Select
  
  CompruebaNumero = intTecla
End Function

'Comprueba que la tecla pulsada corresponde a un nº, decimal o no
Public Function CompruebaNumeroConSigno(intTecla As Integer, strTexto As String) As Integer
  Select Case intTecla
    Case Asc("0") To Asc("9")
      'Se ha pulsado un nº
    
    Case Asc("."), Asc(",")
      Select Case intTecla
        'Case (Asc(".") And FormatoNumericoOk)
          'Se ha pulsado el punto y es el separador decimal correcto
        
        Case (Asc(".") And Not FormatoNumericoOk)
          'Se ha pulsado el punto y NO es el separador decimal correcto
          'Pasar a coma si no hay otra
          intTecla = Asc(",")
        
        'Case Asc(",") And Not FormatoNumericoOk
          'Se ha pulsado la coma y es el separador decimal correcto
        
        Case (Asc(",") And FormatoNumericoOk)
          'Se ha pulsado la coma y NO es el separador decimal correcto
          'Pasar a punto si no hay otro
          intTecla = Asc(".")
      
      End Select
      
      If InStr(strTexto, Chr(intTecla)) > 0 Then
        'Ya existe un punto en la cadena
        intTecla = 0
        Beep
      End If
      
      
    Case 8, 45
      'Se ha pulsado BCKSPC o -
    Case Else
      'Si no son nºs, el punto o BCKSPC, no se admite
      intTecla = 0
      Beep
  End Select
  
  CompruebaNumeroConSigno = intTecla
End Function

'Conjunto de teclas válidas alfanuméricas
Public Function PulsaTecla(intKeyAsc As Integer) As Integer
  Select Case intKeyAsc
    Case 8
      'Backspace
    Case 13
      'Intro
      'SendTab
    Case 32 To 126
      'Caracteres Normales
    Case 225, 233, 237, 243, 250, 241
      'Vocales en minúsculas con acento y ñ minuscula
    Case 193, 201, 205, 211, 218, 209
      'Vocales en Mayúsculas con acento y ñ mayúsculas
    Case 145, 146, 161, 166, 168, 170, 176, 180, 191
      'Símbolos especiales ' ' ¡ | " ª º ´ ¿
    Case Else
      intKeyAsc = 0
      Beep
  End Select
  
  PulsaTecla = intKeyAsc
End Function

'Comprueba que el número pasado no tenga más enteros ni más decimales
'de los especificados por intNumEnteros y intNumDecimales. Devuelve True si el
'formato es correcto o false en caso contrario
Public Function CompruebaFormatoNumero(strTxtTexto As String, intNumEnteros As Integer, intNumDecimales As Integer) As Boolean
  Dim strDecimales As String
  Dim strEnteros As String
  Dim intPunto As Integer
  
  'Situa el lugar del punto decimal
  intPunto = InStr(Format(strTxtTexto, "General Number"), ".")
  CompruebaFormatoNumero = True
  
  If intPunto Then
    'Es un nº con decimales
    strEnteros = Left(strTxtTexto, intPunto - 1)
    strDecimales = (Mid(strTxtTexto, intPunto + 1))
    If Len(strEnteros) > intNumEnteros Then
      MsgBox "La parte entera excede de " & intNumEnteros & " caracteres.", vbOKOnly + vbExclamation
      CompruebaFormatoNumero = False
    End If
    If Len(strDecimales) > intNumDecimales Then
      MsgBox "La parte decimal excede de " & intNumDecimales & " caracteres.", vbOKOnly + vbExclamation
      CompruebaFormatoNumero = False
    End If
  Else
    'Solo hay enteros
    strEnteros = Format(strTxtTexto, "General Number")
    If Len(strEnteros) > intNumEnteros Then
      MsgBox "La parte entera excede de " & intNumEnteros & " caracteres.", vbOKOnly + vbExclamation
      CompruebaFormatoNumero = False
    End If
  End If
End Function

'Comprueba si el formato numérico es "." para decimal, "," para miles
'y "-" para los negativos
Public Function FormatoNumericoOk() As Boolean
  Dim strSeparador As String
  Dim strMiles As String
  Dim strNegativo As String

  strMiles = Right(Left(Format("1000", "#,###"), 2), 1)
  strSeparador = Left(Format("0.50", "#.00"), 1)
  strNegativo = Left(Format("-1", "##"), 1)

  FormatoNumericoOk = IIf(strMiles = "," And strSeparador = "." And strNegativo = "-", True, False)
End Function

'Esta función comprueba que el array de TextBox pasado no contenga ningún
'campo vacío. Se pasa el contador para mostrar la etiqueta asociada si
'fuera necesario
Public Function CamposVacios(objTextoTxt As Object, Optional Cont As Integer, Optional intPrimerIndice As Integer, Optional intUltimoIndice As Integer) As Boolean
  Dim intIndiceBajo As Integer
  Dim intIndiceAlto As Integer
  
  CamposVacios = False
  
  If intPrimerIndice <> 0 And intUltimoIndice <> 0 Then
    'Desde el 1º valor pasado hasta el último valor pasado
    intIndiceBajo = intPrimerIndice
    intIndiceAlto = intUltimoIndice
  Else
    'Los demás casos posibles
    If intPrimerIndice = 0 And intUltimoIndice = 0 Then
      'Desde el 1º hasta el último
      intIndiceBajo = objTextoTxt.LBound
      intIndiceAlto = objTextoTxt.UBound
    ElseIf intPrimerIndice = 0 Then
      'Desde el 1º hasta el valor pasado como último
      intIndiceBajo = objTextoTxt.LBound
      intIndiceAlto = intUltimoIndice
    Else
      'Desde el valor pasado como 1º hasta el último
      intIndiceBajo = intPrimerIndice
      intIndiceAlto = objTextoTxt.UBound
    End If
  End If
  
  For Cont = intIndiceBajo To intIndiceAlto
    If objTextoTxt(Cont) = "" Then
      CamposVacios = True
      Exit Function
    End If
  Next
End Function

'Esta función vacía todos los elementos de una matriz de TextBox
Public Sub VaciarCampos(objTextoTxt As Object)
  Dim I As Integer

  For I = objTextoTxt.LBound To objTextoTxt.UBound
    objTextoTxt(I).Text = ""
  Next
End Sub

'Esta función comprueba que el array de TextBox pasado no contenga ninguna
'fecha incorrecta. La cadena contiene los índices en los que hay que comprobar la fecha
Public Function FechasIncorrectas(objTextoTxt As Object, strCadena As String, Optional intPrimerIndice As Integer, Optional intUltimoIndice As Integer)
  Dim Cont As Integer
  Dim intIndiceBajo As Integer, intIndiceAlto As Integer

  FechasIncorrectas = False

  If intPrimerIndice <> 0 And intUltimoIndice <> 0 Then
    'Desde el 1º valor pasado hasta el último valor pasado
    intIndiceBajo = intPrimerIndice
    intIndiceAlto = intUltimoIndice
  Else
    If intPrimerIndice = 0 And intUltimoIndice = 0 Then
      'Desde el 1º hasta el último
      intIndiceBajo = objTextoTxt.LBound
      intIndiceAlto = objTextoTxt.UBound
    ElseIf intPrimerIndice = 0 Then
      'Desde el 1º hasta el valor pasado como último
      intIndiceBajo = objTextoTxt.LBound
      intIndiceAlto = intUltimoIndice
    Else
      'Desde el valor pasado como 1º hasta el último
      intIndiceBajo = intPrimerIndice
      intIndiceAlto = objTextoTxt.UBound
    End If
  End If

  For Cont = intIndiceBajo To intIndiceAlto
    'Buscar el índice actual en la cadena de índices para ver si fue pasado
    If InStr(strCadena, Cont) > 0 Then
      If Not IsDate(objTextoTxt(Cont).Text) Then
        'Se encontró un campo fecha incorrecto
        FechasIncorrectas = True
        MsgBox "Esa fecha es incorrecta."
        Exit Function
      End If
    End If
  Next
End Function

'Selecciona todos los datos del campo pasado
Public Sub SeleccionaTexto(TxtTexto As Object)
  With TxtTexto
    .SelStart = 0
    .SelLength = Len(TxtTexto)
  End With
End Sub
