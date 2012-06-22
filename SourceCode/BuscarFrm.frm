VERSION 5.00
Begin VB.Form BuscarFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda de un Registro"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   Icon            =   "BuscarFrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   6540
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraOp 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opciones"
      Height          =   690
      Left            =   1500
      TabIndex        =   17
      Top             =   1200
      Width           =   4815
      Begin VB.OptionButton Opt2 
         BackColor       =   &H00FFEED9&
         Caption         =   "Distinto"
         Height          =   195
         Index           =   2
         Left            =   3480
         TabIndex        =   4
         Top             =   300
         Width           =   855
      End
      Begin VB.OptionButton Opt2 
         BackColor       =   &H00FFEED9&
         Caption         =   "Contiene"
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   2
         Top             =   300
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Opt2 
         BackColor       =   &H00FFEED9&
         Caption         =   "Igual"
         Height          =   195
         Index           =   1
         Left            =   2100
         TabIndex        =   3
         Top             =   300
         Width           =   675
      End
   End
   Begin VB.Frame FraOpciones 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opciones para Fechas"
      Height          =   690
      Left            =   1500
      TabIndex        =   16
      Top             =   1200
      Visible         =   0   'False
      Width           =   4815
      Begin VB.OptionButton Opt 
         BackColor       =   &H00FFEED9&
         Caption         =   "<>"
         Height          =   195
         Index           =   5
         Left            =   3900
         TabIndex        =   10
         Top             =   300
         Width           =   615
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H00FFEED9&
         Caption         =   ">"
         Height          =   195
         Index           =   4
         Left            =   3180
         TabIndex        =   9
         Top             =   300
         Width           =   615
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H00FFEED9&
         Caption         =   ">="
         Height          =   195
         Index           =   3
         Left            =   2460
         TabIndex        =   8
         Top             =   300
         Width           =   615
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H00FFEED9&
         Caption         =   "="
         Height          =   195
         Index           =   2
         Left            =   1740
         TabIndex        =   7
         Top             =   300
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H00FFEED9&
         Caption         =   "<="
         Height          =   195
         Index           =   1
         Left            =   1020
         TabIndex        =   6
         Top             =   300
         Width           =   615
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H00FFEED9&
         Caption         =   "<"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   5
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.CommandButton BtnCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3300
      TabIndex        =   12
      Top             =   2025
      Width           =   3165
   End
   Begin VB.CommandButton BtnAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   75
      TabIndex        =   11
      Top             =   2025
      Width           =   3165
   End
   Begin VB.ComboBox CmbNombre 
      BackColor       =   &H00E6FFFF&
      Height          =   315
      Left            =   1500
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   450
      Width           =   3120
   End
   Begin VB.TextBox TxtDato 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1500
      MaxLength       =   50
      TabIndex        =   1
      Top             =   840
      Width           =   4815
   End
   Begin VB.ComboBox CmbCampo 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4125
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   525
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Lbls 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Presione ""Cancelar"" para ver todos los registros"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   18
      Top             =   120
      Width           =   6435
   End
   Begin VB.Label Lbls 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Campo a Buscar"
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   14
      Top             =   510
      Width           =   1170
   End
   Begin VB.Label Lbls 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Valor Buscado"
      Height          =   195
      Index           =   1
      Left            =   285
      TabIndex        =   13
      Top             =   885
      Width           =   1035
   End
End
Attribute VB_Name = "BuscarFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Este formulario sirve para buscar en cualquier form maestro que tenga un grid
'Basa su funcionamiento en leer los nombres del grid pasado y los nombres del campo en la BD
'Evidentemente, sólo se podrá buscar por un campo que además debe estar incluído en el grid
'Puesto que no se puede diferenciar entre cadenas, números y fechas, nunca se podrán
'buscar nºs decimales (porque entre ' ' da error)

'De aquí se cogerán los valores, strAnd indica si se debe de hacer un "AND" o un "WHERE"
Public mDBGrid As DataGrid, mAnd As String
'Esta variable contendrá la cadena terminada con el where hecho
Public mSqlFinal As String

'Guarda el tipo de opción seleccionada para los campos fecha
Dim mstrOpcion As String

'Se carga el form
Private Sub Form_Load()
  Dim I As Integer
  
  CmbNombre.Clear
  CmbCampo.Clear
  
  'Recorrer las columnas del Grid y rellenar la combo con ellas
  For I = 0 To mDBGrid.Columns.Count - 1
    CmbNombre.AddItem mDBGrid.Columns(I).Caption
    CmbCampo.AddItem mDBGrid.Columns(I).DataField
  Next I
  
  mstrOpcion = "="
End Sub

'Se activa el form
Private Sub Form_Activate()
  'Si sólo hay una opción, ponerla automáticamente
  If CmbNombre.ListCount = 1 Then
    CmbNombre.ListIndex = 0
    If TxtDato.Enabled And TxtDato.Visible Then TxtDato.SetFocus
  End If
End Sub

'Crear cadena Slq
Private Sub BtnAceptar_Click()
  If CmbNombre <> "" And TxtDato <> "" Then
    TxtDato = TxtASql(TxtDato)
    
    'Para los campos fecha en access, se emplea # no '
    If UCase(Left(CmbCampo, 5)) = "FECHA" Then
      'Para Access la 2ª o 3ª opción
      mSqlFinal = mAnd & " " & CmbCampo & " " & mstrOpcion & " DateValue('" & Format(TxtDato, "dd/mm/yyyy") & "')"
    Else
      mSqlFinal = mAnd & " " & CmbCampo
      
      'Si la opción es Like, añadirle al dato los %
      If Opt2(0).Value Then TxtDato = "%" & TxtDato & "%"
      
      If Opt2(0).Value Then
        mSqlFinal = mSqlFinal & " Like '" & TxtDato & "'"
      ElseIf Opt2(1).Value Then
        mSqlFinal = mSqlFinal & " = '" & TxtDato & "'"
      ElseIf Opt2(2).Value Then
        mSqlFinal = mSqlFinal & " <> '" & TxtDato & "'"
      End If
      
    End If
    Me.Hide
  Else
    MsgBox "Rellene todos los campos"
  End If
End Sub

'Salir
Private Sub BtnCancelar_Click()
  mSqlFinal = ""
  Me.Hide
End Sub

'Selecciona el texto de los campos TxtDto al seleccionar su TextBox
Private Sub TxtDato_GotFocus()
  SeleccionaTexto TxtDato
End Sub

'Validación de un TextBox antes de salir
Private Sub TxtDato_Validate(Cancel As Boolean)
  If UCase(Left(CmbCampo, 5)) = "FECHA" Then
    If TxtDato <> "" Then
      If Not IsDate(TxtDato) Then
        MsgBox "Esa fecha no es válida."
        Cancel = True
      Else
        TxtDato = Format(TxtDato, "dd/mm/yyyy")
      End If
    End If
  End If
End Sub

'Seleccionar también el Campo de BD
Private Sub CmbNombre_Click()
  If CmbNombre <> "" Then
    CmbCampo.ListIndex = CmbNombre.ListIndex
    
    'Si el campo es fecha visualizar las opciones especiales de fecha
    If UCase(Left(CmbCampo, 5)) = "FECHA" Then
      FraOp.Visible = False
      FraOpciones.Visible = True
    Else
      FraOp.Visible = True
      FraOpciones.Visible = False
    End If
  Else
    CmbCampo.ListIndex = -1
  End If
End Sub

'Seleccionar la opción de fechas
Private Sub Opt_Click(Index As Integer)
  mstrOpcion = Opt(Index).Caption
End Sub
