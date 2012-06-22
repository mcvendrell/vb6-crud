Attribute VB_Name = "CRUDMod"
Option Explicit

'Este módulo contiene funciones que se emplearán en TODOS los formularios
'de MANTENIMIENTO que sigan la norma ESTÁNDAR, esto es, con opciones
'de Cerrar, Nuevo, Editar, Borrar, Grabar, Anular y Buscar, y se limiten
'al mantenimiento de una sola tabla maestra con esos procesos

'Esta función configura un DBGrid pasado con los valores predeterminados
'Si se pasa la opción blnNoSelecRow, la fila actual no se muestra (útil para grids con multiselect)
Public Sub ConfiguraDBGridForm(Grid As DataGrid, Optional blnNoSelecRow As Boolean)
  Dim I As Integer
  Dim strCampo As String
  
  Grid.ScrollBars = dbgVertical
  Grid.Splits(0).Locked = True
  
  If blnNoSelecRow Then
    Grid.MarqueeStyle = dbgNoMarquee
  Else
    Grid.MarqueeStyle = dbgHighlightRow
  End If
  
  'Alinear automáticamente los campos según su tipo
  For I = 0 To Grid.Columns.Count - 1
    strCampo = Grid.Columns(I).DataField
    'Si es un campo normal y no tiene "_", dará error
    If InStr(1, strCampo, "_") Then
      strCampo = Mid(Grid.Columns(I).DataField, 1, InStr(1, Grid.Columns(I).DataField, "_") - 1)
    End If
  
    Select Case strCampo
      Case "NUM", "COD", "CANTIDAD", "IMPORTE", "PRECIO", "TOTAL"
        Grid.Splits(0).Columns(I).Alignment = dbgRight
      Case "FECHA"
        Grid.Splits(0).Columns(I).Alignment = dbgCenter
      Case Else
        Grid.Splits(0).Columns(I).Alignment = dbgLeft
        Grid.Splits(0).Columns(I).Alignment = dbgLeft
    End Select
  Next I
End Sub

'Este proceso inicializa los controles de un form pasado
Public Sub InicializaControlesForm(frmForm As Object)
  Dim Ctrl As Control
  
  For Each Ctrl In frmForm.Controls
    If TypeOf Ctrl Is TextBox Then Ctrl.Text = ""
    If TypeOf Ctrl Is ComboBox Or TypeOf Ctrl Is ListBox Then Ctrl.ListIndex = -1
    If TypeOf Ctrl Is CheckBox Then Ctrl.Value = Unchecked
  Next
End Sub

'Proceso que vacía los campos y los prepara para la introducción de datos
'Configura los botones por defecto
Public Sub InicializaCamposForm(frmForm As Object, blnBotonesTambien As Boolean)
  InicializaControlesForm frmForm
  
  'Posicionar el cursor en el primer campo si es posible
  If frmForm.TxtOb(0).Visible And frmForm.TxtOb(0).Enabled Then frmForm.TxtOb(0).SetFocus
  
  If blnBotonesTambien Then ConfiguraBotonesForm frmForm, False
  
  'En este modo no estamos modificando
  frmForm.blnModificando = False
End Sub

'Este proceso habilita/inhabilita los controles de un form ESTANDAR pasado
Public Sub ConfiguraControlesForm(frmForm As Object, blnModoLectura As Boolean)
  Dim Ctrl As Control
  
  For Each Ctrl In frmForm.Controls
    If TypeOf Ctrl Is TextBox Or _
       TypeOf Ctrl Is ComboBox Or _
       TypeOf Ctrl Is ListBox Or _
       TypeOf Ctrl Is CheckBox Or _
       TypeOf Ctrl Is OptionButton Then Ctrl.Enabled = Not blnModoLectura
  Next

  'Si se está en modo edición, no permitir la modificación de la clave ni tocar el grid
  'En modo inserción no tocar el grid
  If frmForm.blnModificando Then
    frmForm.TxtOb(0).Enabled = False
    frmForm.DBGridDatos.Enabled = False
  ElseIf frmForm.blnInsertando Then
    frmForm.DBGridDatos.Enabled = False
  Else
    frmForm.DBGridDatos.Enabled = True
  End If
End Sub

'Llena los controles de un form ESTANDAR con los datos de un Resultset pasado
Public Sub RefrescaCamposForm(frmForm As Object, RecDatos As Recordset)
  Dim Ctrl As Control
 
  With RecDatos
    If Not .BOF And Not .EOF Then
      'Existen datos
      On Local Error GoTo Errores
      For Each Ctrl In frmForm.Controls
        If TypeOf Ctrl Is TextBox Then
          'Muestra los valores de Txts, Masks y NumberMasks
          'Si no hay tag no se rellena el combo
          '(para los txt internos y ocultos)
          If Ctrl.Tag <> "" Then Ctrl.Text = "" & .Fields(Ctrl.Tag)
        End If
        
        If TypeOf Ctrl Is ComboBox Then
          'Muestra los valores de las combos
          'Si no hay tag no se rellena el combo
          '(para la pareja de combos Cod-Descripción)
          If Ctrl.Tag <> "" Then
            If IsNull(.Fields(Ctrl.Tag)) Then
              Ctrl.ListIndex = -1
            Else
              Ctrl.Text = .Fields(Ctrl.Tag)
            End If
          End If
        End If
        
        If TypeOf Ctrl Is CheckBox Then
          If Ctrl.Tag <> "" Then
            'Muestra los valores de los checks
            Ctrl.Value = IIf(.Fields(Ctrl.Tag) = "S", Checked, Unchecked)
          End If
        End If
      
        If TypeOf Ctrl Is OptionButton Then
          If Ctrl.Tag <> "" Then
            'Muestra los valores de los option
            'El TAG está compuesto por NOMBRE_CAMPO=Valor que corresponde al Option
            'Ejemplo ESTADO=B. Si El Ctrl actual tiene como Valor el valor grabado en
            'el campo NOMBRE_CAMPO de la BD, se pone a true
            Ctrl.Value = IIf(.Fields(Left(Ctrl.Tag, Len(Ctrl.Tag) - 2)) = Right(Ctrl.Tag, 1), True, False)
          End If
        End If
      
      Next Ctrl
      
      On Local Error GoTo ErroresForm
      ConfiguraBotonesForm frmForm, True
      ConfiguraControlesForm frmForm, True

      'Marcar la posición actual para recuperarla después
      frmForm.Marca = .Bookmark
    End If
  End With
  
  'Cuando se refrescan los campos, ni estamos en modo insertar, ni modificar
  frmForm.blnInsertando = False
  frmForm.blnModificando = False
  
  If frmForm.DBGridDatos.Enabled And frmForm.DBGridDatos.Visible Then frmForm.DBGridDatos.SetFocus
  Exit Sub
  
Errores:
  If Ctrl.Tag <> "" Then
    If Err.Number = 3265 Then
      MsgBox "Error en el TAG del campo '" & Ctrl.Name & "(" & Ctrl.Index & ")' valor -> " & Ctrl.Tag & ": " & Err.Description
    Else
      MsgBox "Error en un campo (" & Ctrl.Name & " valor -> " & RecDatos.Fields(Ctrl.Tag) & "): " & Err.Description
    End If
  Else
    MsgBox "Error en el campo '" & Ctrl.Name & "': no tiene TAG"
  End If
  Resume Next
ErroresForm:
  MsgBox "Error en un proceso del formulario: " & Err.Description
  Resume Next
End Sub

'Refresca el Rec y el grid de datos (que está asociado)
Public Sub RefrescaGridForm(frmForm As Object, RecDatos As Recordset)
  RecDatos.Requery
  RecuperaMarcaForm frmForm, RecDatos
End Sub

'Restablece la posición en la que se estaba
Public Sub RecuperaMarcaForm(frmForm As Object, RecDatos As Recordset)
  On Local Error GoTo Errores
  With frmForm
    If Not RecDatos.EOF And Not RecDatos.BOF Then
      If frmForm.Marca > 0 Then
        RecDatos.Bookmark = frmForm.Marca
      Else
        RecDatos.MoveFirst
      End If
      RefrescaCamposForm frmForm, RecDatos
      frmForm.LblDatos.Caption = " Registro: " & (RecDatos.AbsolutePosition) & " de " & RecDatos.RecordCount
    Else
      'No hay datos
      InicializaCamposForm frmForm, False
      frmForm.LblDatos.Caption = " No hay datos"
    End If
  End With
Errores:
  'No hacer nada
End Sub

'Acciones cuando se cambia de fila en un grid
Public Sub CambiaFilaForm(frmForm As Object, RecDatos As Recordset)
  Dim Cont As Integer
  
  On Local Error GoTo Errores
  If Not RecDatos.BOF And Not RecDatos.EOF Then
    'Refresca los datos
    RefrescaCamposForm frmForm, RecDatos
    frmForm.LblDatos.Caption = " Registro: " & (RecDatos.AbsolutePosition) & " de " & RecDatos.RecordCount
  Else
    'Si no hay registros se deshabilitan los botones de Editar y Borrar
    frmForm.BarMenu.Buttons("Editar").Enabled = False
    frmForm.BarMenu.Buttons("Borrar").Enabled = False
   
    'Y deshabilitamos los botones de avance y retroceso de los registros en el grid
    For Cont = 0 To 3
      frmForm.BtnDatos(Cont).Enabled = False
    Next Cont
  End If
Errores:
  'No hacer nada
End Sub

'Proceso que habilita los botones pertinentes según estemos insertando o modificando
Public Sub ConfiguraBotonesForm(frmForm As Object, blnModoLectura As Boolean)
  Dim I As Integer
  
  With frmForm.BarMenu
    .Buttons("Cerrar").Enabled = blnModoLectura
    .Buttons("Nuevo").Enabled = blnModoLectura
    .Buttons("Editar").Enabled = blnModoLectura
    .Buttons("Borrar").Enabled = blnModoLectura
    .Buttons("Grabar").Enabled = Not blnModoLectura
    .Buttons("Anular").Enabled = Not blnModoLectura
    .Buttons("Buscar").Enabled = blnModoLectura
  End With
  
  'Botones de control de datos
  For I = 0 To 3
    frmForm.BtnDatos(I).Enabled = blnModoLectura
  Next
End Sub

'Control de la tecla pulsada en un form
Public Sub PulsaTeclaForm(frmForm As Object, intTecla As Integer, Shift As Integer)
  On Local Error Resume Next
  
  With frmForm
    'Si se está insertando o modificando, no hacer nada
    If .blnInsertando Or .blnModificando Then
      'Salvo si la tecla es Esc, que debe anular la operación,
      'o ENTER que debe grabar
      If intTecla = vbKeyEscape Then
        If .BarMenu.Buttons("Anular").Enabled Then Call .barMenu_ButtonClick(.BarMenu.Buttons("Anular"))
      ElseIf intTecla = vbKeyReturn Then
        If .BarMenu.Buttons("Grabar").Enabled Then Call .barMenu_ButtonClick(.BarMenu.Buttons("Grabar"))
      End If
      
    Else
    
      'Modo normal
      Select Case intTecla
        Case vbKeyEscape, vbKeyC
          If .BarMenu.Buttons("Cerrar").Enabled Then CerrarForm frmForm
        'Botones de la barra
        Case vbKeyN
          If .BarMenu.Buttons("Nuevo").Enabled Then Call .barMenu_ButtonClick(.BarMenu.Buttons("Nuevo"))
        Case vbKeyE
          If .BarMenu.Buttons("Editar").Enabled Then Call .barMenu_ButtonClick(.BarMenu.Buttons("Editar"))
        Case vbKeyB
          If .BarMenu.Buttons("Borrar").Enabled Then Call .barMenu_ButtonClick(.BarMenu.Buttons("Borrar"))
        Case vbKeyU
          If .BarMenu.Buttons("Buscar").Enabled Then Call .barMenu_ButtonClick(.BarMenu.Buttons("Buscar"))
        'Acciones especiales sobre el grid
        Case vbKeyEnd
          If .BtnDatos_Click(3).Enabled Then Call .BtnDatos_Click(3)
        Case vbKeyHome
          If .BtnDatos_Click(0).Enabled Then Call .BtnDatos_Click(0)
        Case vbKeyUp, vbKeyPageUp
          If Shift = vbCtrlMask Then
            If .BtnDatos_Click(0).Enabled Then Call .BtnDatos_Click(0)
          Else
            If .BtnDatos_Click(1).Enabled Then Call .BtnDatos_Click(1)
          End If
        Case vbKeyDown, vbKeyPageDown
          If Shift = vbCtrlMask Then
            If .BtnDatos_Click(3).Enabled Then Call .BtnDatos_Click(3)
          Else
            If .BtnDatos_Click(2).Enabled Then Call .BtnDatos_Click(2)
          End If
      End Select
    End If
  End With
End Sub

'El usuario quiere grabar un nuevo registro
Public Sub NuevoForm(frmForm As Object)
  With frmForm
    frmForm.blnInsertando = True
    frmForm.LblDatos.Caption = " Insertando un registro"
    ConfiguraControlesForm frmForm, False
    InicializaCamposForm frmForm, True
  End With
End Sub

'El usuario quiere ponerse en modo edición del registro actual
Public Sub EditarForm(frmForm As Object)
  With frmForm
    frmForm.blnModificando = True
    frmForm.LblDatos.Caption = " Modificando el registro"
    ConfiguraControlesForm frmForm, False
    ConfiguraBotonesForm frmForm, False
    
    'Deshabilitar el campo clave para edición
    frmForm.TxtOb(0).Enabled = False
    
    On Local Error Resume Next
    If frmForm.TxtOb(1).Visible Then frmForm.TxtOb(1).SetFocus
  End With
End Sub

'El usuario quiere borrar el registro seleccionado
Public Sub BorrarForm(frmForm As Object, Conexion As ADODB.Connection, Tabla As String, RecDatos As Recordset, strMensaje As String, Optional blnIDNumerico As Boolean)
  Dim Sql As String
  
  On Local Error GoTo Errores
  
  If MsgBox(strMensaje, vbYesNo + vbQuestion + vbDefaultButton2, "Borrar datos") = vbYes Then
    Sql = "delete from " & Tabla _
        & " where " & frmForm.TxtOb(0).Tag
    If blnIDNumerico Then
      Sql = Sql & " = " & frmForm.TxtOb(0)
    Else
      Sql = Sql & " = '" & frmForm.TxtOb(0) & "'"
    End If
    Conexion.Execute Sql
    
    RefrescaGridForm frmForm, RecDatos
  End If
  Exit Sub

Errores:
  MsgBox "Se produjo un error: " & Err.Description
End Sub

'El usuario ha solicitado la cancelación de nuevo/editar registro
Public Sub AnularForm(frmForm As Object, RecDatos As Recordset)
  With frmForm
    .blnInsertando = False
    .blnModificando = False
    ConfiguraControlesForm frmForm, True
    ConfiguraBotonesForm frmForm, True
    
    RecuperaMarcaForm frmForm, RecDatos
  End With
End Sub

'Oculta el formulario y vuelve a la pantalla principal
Public Sub CerrarForm(frmForm As Object)
  Unload frmForm
End Sub

'Proceso que mueve un registro hasta la primera posición
Public Sub MuevePrimero(Rec As Recordset)
  On Local Error GoTo Error
  Rec.MoveFirst
  Exit Sub
Error:
  MsgBox "Imposible moverse al primer registro"
End Sub

'Proceso que mueve un registro hasta la anterior posición
Public Sub MueveAnterior(Rec As Recordset)
  On Local Error GoTo Error
  If Not Rec.BOF Then Rec.MovePrevious
  If Rec.BOF Then Rec.MoveFirst
  Exit Sub
Error:
  MsgBox "Imposible moverse al anterior registro"
End Sub

'Proceso que mueve un registro hasta la siguiente posición
Public Sub MueveSiguiente(Rec As Recordset)
  On Local Error GoTo Error
  If Not Rec.EOF Then Rec.MoveNext
  If Rec.EOF Then Rec.MoveLast
  Exit Sub
Error:
  MsgBox "Imposible moverse al siguiente registro"
End Sub

'Proceso que mueve un registro hasta la última posición
Public Sub MueveUltimo(Rec As Recordset)
  On Local Error GoTo Error
  Rec.MoveLast
  Exit Sub
Error:
  MsgBox "Imposible moverse al último registro"
End Sub

'Esta función devuelve el siguiente código disponible para una tabla pasada
'si los códigos son consecutivos y numéricos.
'Se puede restringir la búsqueda (por si la clave es múltiple)
Public Function NuevoCodigo(ByRef Conexion As ADODB.Connection, ByRef strCampo As String, ByRef strTabla As String, Optional strRestriccion As String) As Long
  Dim Sql As String, Rec As New Recordset
  
  On Local Error GoTo Errores
  
  Sql = "select Max(" & strCampo & ") As NUM" _
      & "  from " & strTabla
  If strRestriccion <> "" Then Sql = Sql & " where " & strRestriccion
  Rec.Open Sql, Conexion
  
  If Not Rec.EOF Then
    If Not IsNull(Rec!NUM) Then
      NuevoCodigo = Rec!NUM + 1
    Else
      NuevoCodigo = 1
    End If
  Else
    NuevoCodigo = 1
  End If
  
  Rec.Close
  Set Rec = Nothing
  Exit Function

Errores:
  MsgBox "Error intentando calcular el siguiente código: " & Err.Description
  Set Rec = Nothing
End Function

'Esta función busca un código alfanumérico pasado en una tabla y campo pasados.
'Si se pasan restricciones, se aplican en la busqueda
Public Function Busca(ByRef Conexion As ADODB.Connection, ByRef strCod As String, ByRef strCampo As String, ByRef strTabla As String, Optional blnNumerico As Boolean, Optional strRestriccion As String) As Boolean
  Dim Rec As New Recordset, Sql As String
  
  On Local Error GoTo Errores
  
  'Se seleccionan los registros que coincidan con el código pasado
  Sql = "select " & strCampo _
      & "  from " & strTabla _
      & " where " & strCampo & " = " _
      & IIf(blnNumerico, strCod, "'" & strCod & "'") _
      & IIf(strRestriccion = "", "", " and " & strRestriccion)
  Rec.Open Sql, Conexion
  
  If Rec.BOF And Rec.EOF Then
    'No hay registros
    Busca = False
  Else
    Busca = True
  End If
  
  Rec.Close
  Set Rec = Nothing
  Exit Function

Errores:
  MsgBox "Error en búsqueda: " & Err.Description
  Set Rec = Nothing
End Function
