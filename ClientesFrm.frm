VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ClientesFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de CLIENTES"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   825
   ClientWidth     =   9345
   Icon            =   "ClientesFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   9345
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkBaja 
      BackColor       =   &H00FFEED9&
      Caption         =   "Baja"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7080
      TabIndex        =   37
      Tag             =   "BAJA"
      Top             =   750
      Width           =   675
   End
   Begin VB.TextBox TxtOb 
      BackColor       =   &H00E6FFFF&
      Height          =   285
      Index           =   2
      Left            =   1020
      MaxLength       =   9
      TabIndex        =   4
      Tag             =   "CIF"
      Top             =   1380
      Width           =   1020
   End
   Begin VB.TextBox TxtPres 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   555
      Index           =   0
      Left            =   1020
      MaxLength       =   2000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Tag             =   "COMENTARIOS"
      Top             =   2805
      Width           =   8280
   End
   Begin VB.TextBox TxtOb 
      BackColor       =   &H00E6FFFF&
      Height          =   285
      Index           =   1
      Left            =   1020
      MaxLength       =   60
      TabIndex        =   2
      Tag             =   "APELLIDOS"
      Top             =   1050
      Width           =   3780
   End
   Begin VB.TextBox TxtPres 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   7
      Left            =   1020
      MaxLength       =   40
      TabIndex        =   12
      Tag             =   "MAIL"
      Top             =   2460
      Width           =   3890
   End
   Begin VB.ComboBox CmbLocalidad 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Tag             =   "LOCALIDAD"
      ToolTipText     =   "Para vaciar la casilla hacer doble click sobre el nombre"
      Top             =   1725
      Width           =   2850
   End
   Begin VB.TextBox TxtPres 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   1020
      MaxLength       =   15
      TabIndex        =   9
      Tag             =   "TLFN1"
      Top             =   2115
      Width           =   1680
   End
   Begin VB.TextBox TxtOb 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   0
      Left            =   1020
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "COD_CLIENTE"
      Top             =   705
      Width           =   660
   End
   Begin VB.TextBox TxtPres 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   5610
      MaxLength       =   10
      TabIndex        =   19
      Tag             =   "FECHA_BAJA"
      Top             =   705
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.ComboBox CmbProvincia 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   6630
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Tag             =   "PROVINCIA"
      ToolTipText     =   "Para vaciar la casilla hacer doble click sobre el nombre"
      Top             =   1725
      Width           =   2670
   End
   Begin VB.TextBox TxtPres 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   5
      Left            =   3225
      MaxLength       =   15
      TabIndex        =   10
      Tag             =   "TLFN2"
      Top             =   2115
      Width           =   1680
   End
   Begin VB.TextBox TxtPres 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   3150
      MaxLength       =   80
      TabIndex        =   5
      Tag             =   "DIRECCION"
      Top             =   1380
      Width           =   6150
   End
   Begin VB.TextBox TxtPres 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   5055
      MaxLength       =   5
      TabIndex        =   7
      Tag             =   "COD_POSTAL"
      Top             =   1725
      Width           =   630
   End
   Begin VB.TextBox TxtPres 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   6
      Left            =   5460
      MaxLength       =   15
      TabIndex        =   11
      Tag             =   "FAX"
      Top             =   2115
      Width           =   1680
   End
   Begin VB.TextBox TxtPres 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   8
      Left            =   5610
      MaxLength       =   30
      TabIndex        =   3
      Tag             =   "NOMBRE"
      Top             =   1050
      Width           =   3690
   End
   Begin VB.TextBox TxtOb 
      BackColor       =   &H00E6FFFF&
      Height          =   285
      Index           =   3
      Left            =   3150
      MaxLength       =   10
      TabIndex        =   1
      Tag             =   "FECHA_ALTA"
      Top             =   705
      Width           =   1020
   End
   Begin VB.PictureBox PicRegistro 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   9345
      TabIndex        =   20
      Top             =   5175
      Width           =   9345
      Begin VB.CommandButton BtnDatos 
         Height          =   300
         Index           =   3
         Left            =   8955
         Picture         =   "ClientesFrm.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton BtnDatos 
         Height          =   300
         Index           =   2
         Left            =   8595
         Picture         =   "ClientesFrm.frx":0784
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton BtnDatos 
         Height          =   300
         Index           =   1
         Left            =   360
         Picture         =   "ClientesFrm.frx":0AC6
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton BtnDatos 
         Height          =   300
         Index           =   0
         Left            =   15
         Picture         =   "ClientesFrm.frx":0E08
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label LblDatos 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   765
         TabIndex        =   21
         Top             =   0
         Width           =   7770
      End
   End
   Begin MSComctlLib.Toolbar BarMenu 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   1005
      ButtonWidth     =   1244
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImgBarra"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cerrar"
            Key             =   "Cerrar"
            Description     =   "Cerrar Ventana"
            Object.ToolTipText     =   "Cerrar Ventana"
            ImageIndex      =   1
            Object.Width           =   1000
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Nuevo"
            Key             =   "Nuevo"
            Description     =   "Agregar Registro"
            Object.ToolTipText     =   "Agregar Registro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Editar"
            Key             =   "Editar"
            Description     =   "Modificar Datos"
            Object.ToolTipText     =   "Modificar Datos"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Borrar"
            Key             =   "Borrar"
            Description     =   "Eliminar Registro"
            Object.ToolTipText     =   "Eliminar Registro"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Grabar"
            Key             =   "Grabar"
            Description     =   "Actualizar Cambios"
            Object.ToolTipText     =   "Actualizar Cambios"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Anular"
            Key             =   "Anular"
            Description     =   "Cancelar Cambios"
            Object.ToolTipText     =   "Cancelar Cambios"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "B&uscar"
            Key             =   "Buscar"
            Description     =   "Buscar"
            Object.ToolTipText     =   "Buscar Registro"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImgBarra 
      Left            =   8760
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClientesFrm.frx":114A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClientesFrm.frx":1466
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClientesFrm.frx":1782
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClientesFrm.frx":1A9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClientesFrm.frx":203A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClientesFrm.frx":248E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClientesFrm.frx":27AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClientesFrm.frx":2AC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClientesFrm.frx":2F22
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClientesFrm.frx":337E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClientesFrm.frx":37D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClientesFrm.frx":3C24
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClientesFrm.frx":4076
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DBGridDatos 
      Height          =   1680
      Left            =   60
      TabIndex        =   14
      Top             =   3420
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   2963
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16772825
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Datos Existentes"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "COD_CLIENTE"
         Caption         =   "Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "APELLIDOS"
         Caption         =   "Apellidos"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "NOMBRE"
         Caption         =   "Nombre"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "CIF"
         Caption         =   "C.I.F."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "DIRECCION"
         Caption         =   "Dirección"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "LOCALIDAD"
         Caption         =   "Localidad"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "PROVINCIA"
         Caption         =   "Provincia"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "MAIL"
         Caption         =   "e-mail"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "TLFN1"
         Caption         =   "Tlfn 1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnAllowSizing=   -1  'True
            ColumnWidth     =   629,858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2340,284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1544,882
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   3300,095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2174,74
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2025,071
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2415,118
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1230,236
         EndProperty
      EndProperty
   End
   Begin VB.Label LblPres 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fax"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   5100
      TabIndex        =   36
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label LblPres 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Teléfonos                                             ------"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   75
      TabIndex        =   35
      Top             =   2160
      Width           =   3000
   End
   Begin VB.Label LblOb 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "C.I.F."
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   75
      TabIndex        =   34
      Top             =   1425
      Width           =   375
   End
   Begin VB.Label LblOb 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Código"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   33
      Top             =   750
      Width           =   495
   End
   Begin VB.Label LblPres 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fecha Baja"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   4680
      TabIndex        =   32
      Top             =   750
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label LblPres 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cód.Postal"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   4200
      TabIndex        =   31
      Top             =   1785
      Width           =   765
   End
   Begin VB.Label Lbls 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Provincia"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   5865
      TabIndex        =   30
      Top             =   1785
      Width           =   660
   End
   Begin VB.Label LblPres 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dirección"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   2385
      TabIndex        =   29
      Top             =   1425
      Width           =   675
   End
   Begin VB.Label Lbls 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Localidad"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   28
      Top             =   1785
      Width           =   690
   End
   Begin VB.Label LblPres 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nombre"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   4935
      TabIndex        =   27
      Top             =   1110
      Width           =   555
   End
   Begin VB.Label LblPres 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "e-Mail"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   75
      TabIndex        =   26
      Top             =   2505
      Width           =   420
   End
   Begin VB.Label LblOb 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Apellidos/RS"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   75
      TabIndex        =   25
      Top             =   1110
      Width           =   930
   End
   Begin VB.Label LblOb 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fecha Alta"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   2295
      TabIndex        =   24
      Top             =   750
      Width           =   765
   End
   Begin VB.Label LblPres 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Comentarios"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   23
      Top             =   2835
      Width           =   870
   End
End
Attribute VB_Name = "ClientesFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------   Constantes del formulario   ----------------------
'Nombre de la tabla maestra (si es una, sino obviar)
'Master table name
Const Tabla = "CLIENTES"

'-------------  Variables comunes para todos los mantenimientos  --------
'Variables de cotrol de datos
'Data control vars, always mandatory
Dim Sql As String
Dim RecDatos As New Recordset
Dim mstrOrden As String

'Variables para saber si se está insertando o modificando
'Son públicas para que RefrescaCampos pueda funcionar con el proceso público RefrescaControles
'Public vars to know if we are inserting or editing the form
Public blnInsertando As Boolean
Public blnModificando As Boolean

'Variable que guarda la posición actual en el Recordset para poder recuperarla después.
'Es pública para que RefrescaCampos pueda funcionar con el proceso público RefrescaControles
'Public var that saves the actual Recordset position to recover it after refreshing the data
Public Marca As Variant

'Con DBGrid normal de Microsoft, al hacer un requery de los datos se produce un RowColChange
'que a veces no interesa, se controla con esta variable
'Sometimes we dont want to retrieve all data again, we control it with this var
Dim mblnNoCambiarFila As Boolean

'----------------------   Funciones Generales del form ------------------

'Acciones cuando se cambia de fila en el grid
'Actions for change the row in the grid
Private Sub DBGridDatos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  If Not mblnNoCambiarFila Then CambiaFilaForm Me, RecDatos
End Sub

'Ordenar los datos según se elija
'Ordering data while head clicking
Private Sub DBGridDatos_HeadClick(ByVal ColIndex As Integer)
  Screen.MousePointer = vbHourglass

  'Ordenar por la columna pinchada en orden normal o inverso
  If mstrOrden = DBGridDatos.Columns(ColIndex).DataField Then
    mstrOrden = DBGridDatos.Columns(ColIndex).DataField & " Desc"
  Else
    mstrOrden = DBGridDatos.Columns(ColIndex).DataField
  End If

  'Hacer la selección
  On Local Error Resume Next
  RecDatos.Close
  On Local Error GoTo 0
  Sql = "select * from " & Tabla & " order by " & mstrOrden
  RecDatos.Open Sql, Conexion, adOpenStatic, adLockReadOnly
  Set DBGridDatos.DataSource = RecDatos

  Screen.MousePointer = vbDefault
End Sub

'Se presionó un botón del control de datos
'Actions when we push a buton for the control grid (first-prev-next-last)
Public Sub BtnDatos_Click(Index As Integer)
  Select Case Index
    Case 0
      'Botón de Ir al Primer Registro
      MuevePrimero RecDatos
    Case 1
      'Botón de Ir al Anterior Registro
      MueveAnterior RecDatos
    Case 2
      'Botón de Ir al Siguiente Registro
      MueveSiguiente RecDatos
    Case 3
      'Botón de Ir al Último Registro
      MueveUltimo RecDatos
  End Select
End Sub

'Se carga el formulario
'Loading form actions
Private Sub Form_Load()
  Screen.MousePointer = vbHourglass

  ConfiguraDBGridForm DBGridDatos
  DBGridDatos.ScrollBars = dbgAutomatic

  'Son de ejemplo, se debería definir algún proceso o leer de otra tabla
  'Example fields, you must define your own process or read from other sources
  CmbLocalidad.AddItem "Localidad 1"
  CmbLocalidad.AddItem "Localidad 2"
  CmbLocalidad.AddItem "Localidad 3"
  CmbProvincia.AddItem "Provincia 1"
  CmbProvincia.AddItem "Provincia 2"
  
  'Por defecto estamos en modo consulta
  'Reading by default
  blnInsertando = False
  blnModificando = False
  mblnNoCambiarFila = False
  ConfiguraBotonesForm Me, True
  ConfiguraControlesForm Me, True

  'Fijar las opciones de orden
  'Set your order options
  mstrOrden = "APELLIDOS, NOMBRE"

  'Hacer la selección inicial, asignar el Rec y el Grid
  'Initial select
  Sql = "select * from " & Tabla & " order by " & mstrOrden
  RecDatos.Open Sql, Conexion, adOpenStatic, adLockReadOnly
  Set DBGridDatos.DataSource = RecDatos
  If RecDatos.EOF Then CambiaFilaForm Me, RecDatos

  Screen.MousePointer = vbDefault
End Sub

'Acciones cuando se presion una tecla sobre el form
'Pressing a key
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then
    MsgBox "You can show help here", vbInformation
  Else
    PulsaTeclaForm Me, KeyCode, Shift
  End If
End Sub

'Se descarga el form
'Unloading the form
Private Sub Form_Unload(Cancel As Integer)
  On Local Error Resume Next
  RecDatos.Close
  Set RecDatos = Nothing
End Sub

'Se presionó un botón de la barra de botones
'Actions for pressing the menu bar options
Public Sub barMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case "Cerrar"
      CerrarForm Me
    
    Case "Nuevo"
      NuevoForm Me
    
    Case "Editar"
      EditarForm Me
    
    Case "Borrar"
      BorrarForm Me, Conexion, Tabla, RecDatos, "¿Está seguro de que desea borrar este Registro?", True
    
    Case "Grabar"
      Grabar
    
    Case "Anular"
      AnularForm Me, RecDatos
    
    Case "Buscar"
      Buscar
  
  End Select
End Sub

'El usuario ha confirmado la grabación/modificación de unos datos
'Save the new/edit data
Private Sub Grabar()
  Dim I As Integer

  'Comprobar que no haya campos vacíos
  'Empty fields?
  If CamposVacios(TxtOb, I, 1) Then
    'Algún campo está vacío, Msg de aviso. No se modifica nada
    'Help message because there are mandatory fields not filled
    MsgBox "El campo " & LblOb(I).Caption & " está vacío. Ningún campo amarillo puede quedar vacío.", vbExclamation, "Faltan datos"
    'Se sitúa el cursor en el campo vacío
    'Go to empty field
    TxtOb(I).SetFocus
  Else
    With RecDatos
      On Local Error GoTo Errores

      'Comienza la transacción
      'Make a transaction
      Conexion.BeginTrans

      If blnInsertando Then
        'Se está añadiendo
        'Inserting
        Sql = "insert into " & Tabla _
            & "       (" & TxtOb(0).Tag & "," & TxtOb(1).Tag & "," & TxtOb(2).Tag & "," & TxtOb(3).Tag & ") " _
            & "values (" & TxtOb(0) & ",'" & TxtOb(1) & "','" & TxtOb(2) & "','" & TxtOb(3) & "')"
      Else
        'Se está modificando
        'Editing
        Sql = "update " & Tabla _
            & "   set " & TxtOb(1).Tag & " = '" & TxtOb(1) & "'" _
            & "     , " & TxtOb(2).Tag & " = '" & TxtOb(2) & "'" _
            & "     , " & TxtOb(3).Tag & " = '" & TxtOb(3) & "'" _
            & " where " & TxtOb(0).Tag & " = " & TxtOb(0)
      End If
      Conexion.Execute Sql

      'Campos opcionales
      'Optional fields
      For I = TxtPres.LBound To TxtPres.UBound
        If TxtPres(I) <> "" Then
          Select Case I
            Case 100
              'Campos numéricos (no hay en este caso)
              'Numeric fields (the arent in this case)
              Sql = "update " & Tabla _
                  & "   set " & TxtPres(I).Tag & " = " & TxtPres(I) _
                  & " where " & TxtOb(0).Tag & " = " & TxtOb(0)
            Case Else
              'Campos fecha o caracter
              'Date or string fields
              Sql = "update " & Tabla _
                  & "   set " & TxtPres(I).Tag & " = '" & TxtPres(I) & "'" _
                  & " where " & TxtOb(0).Tag & " = " & TxtOb(0)
          End Select
        Else
          'Campo vacío
          'Empty field
          Sql = "update " & Tabla _
              & "   set " & TxtPres(I).Tag & " = Null" _
              & " where " & TxtOb(0).Tag & " = " & TxtOb(0)
        End If
        Conexion.Execute Sql
      Next

      'Combos
      Sql = "update " & Tabla _
          & "   set " & CmbLocalidad.Tag & " = " & IIf(CmbLocalidad = "", "Null", "'" & CmbLocalidad & "'") _
          & " where " & TxtOb(0).Tag & " = " & TxtOb(0)
      Conexion.Execute Sql

      Sql = "update " & Tabla _
          & "   set " & CmbProvincia.Tag & " = " & IIf(CmbProvincia = "", "Null", "'" & CmbProvincia & "'") _
          & " where " & TxtOb(0).Tag & " = " & TxtOb(0)
      Conexion.Execute Sql

      'Chks
      Sql = "update " & Tabla _
          & "   set " & ChkBaja.Tag & " = '" & IIf(ChkBaja.Value = vbChecked, "S", "N") & "'" _
          & " where " & TxtOb(0).Tag & " = " & TxtOb(0)
      Conexion.Execute Sql

Errores:
      'Si se ha producido un error se deshacen las acciones
      'On error, rollback
      If CompruebaError() Then
        Conexion.RollbackTrans
      Else
        Conexion.CommitTrans

        'Si se ha grabado un registro nuevo, ir a ese registro, sino recuperar consulta
        'When we save a new record, show only this record to let user see it, else, refresh all
        If blnInsertando Then
          On Local Error Resume Next
          RecDatos.Close
          On Local Error GoTo 0
          Sql = "select * from " & Tabla & " where " & TxtOb(0).Tag & " = " & TxtOb(0)
          RecDatos.Open Sql, Conexion, adOpenStatic, adLockReadOnly
          Set DBGridDatos.DataSource = RecDatos
        Else
          mblnNoCambiarFila = True
          RefrescaGridForm Me, RecDatos
          mblnNoCambiarFila = False
          DBGridDatos.Enabled = True
        End If
      End If
    End With
  End If
End Sub

'El usuario ha solicitado la búsqueda de un registro
'Search by field (obtained from DataGrid)
Private Sub Buscar()
  Set BuscarFrm.mDBGrid = DBGridDatos
  BuscarFrm.mAnd = "where"
  BuscarFrm.Show vbModal

  Screen.MousePointer = vbHourglass

  On Local Error Resume Next
  RecDatos.Close
  On Local Error GoTo 0
  If BuscarFrm.mSqlFinal <> "" Then
    Sql = "select * from " & Tabla _
        & " " & BuscarFrm.mSqlFinal _
        & " order by " & mstrOrden
  Else
    Sql = "select * from " & Tabla & " order by " & mstrOrden
  End If
  Unload BuscarFrm
  
  On Error Resume Next
  RecDatos.Open Sql, Conexion, adOpenStatic, adLockReadOnly
  RefrescaGridForm Me, RecDatos

  Screen.MousePointer = vbDefault
End Sub

'Selecciona el texto de los campos TxtOb al seleccionar su TextBox
'Select all text after entering in the field
Private Sub TxtOb_GotFocus(Index As Integer)
  If Index = 0 And blnInsertando Then
    'Buscar el siguiente código, estamos insertando
    'While inserting, give a new code
    TxtOb(Index) = NuevoCodigo(Conexion, TxtOb(Index).Tag, Tabla)
  End If
  SeleccionaTexto TxtOb(Index)
End Sub

'Acciones a realizar cuando se presiona una tecla en un TextBox TxtOb
'Checks while pressing a key in the field
Private Sub TxtOb_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case Index
    Case 0
      'Valida las entradas numéricas
      CompruebaEntero KeyAscii

    Case 3
      'Valida las entradas fecha
      CompruebaFecha KeyAscii

  End Select
End Sub

'Validación de un TextBox TxtOb antes de salir
'Checks while leaving the field
Private Sub TxtOb_Validate(Index As Integer, Cancel As Boolean)
  Select Case Index
    Case 0
      'Comprobar que la clave no sea repetida
      If TxtOb(Index) <> "" And TxtOb(Index).Enabled And TxtOb(Index).Visible Then
        'Se está insertando. Busca una clave repetida
        'Se comprueba Enabled por si se estaba insertando y se pulsó en el
        'control data para ir a otro registro
        If Busca(Conexion, TxtOb(Index), TxtOb(Index).Tag, Tabla, True) Then
          'Se encontró el código, no es nuevo
          MsgBox "Ese código ya existe.", vbInformation
          Cancel = True
        End If
      End If

    Case 2
      'Ver si ya existe ese CIF
      If TxtOb(Index) <> "" Then
        If Busca(Conexion, TxtOb(Index), TxtOb(Index).Tag, Tabla, , TxtOb(0).Tag & " <> " & TxtOb(0)) And blnInsertando Then
          'Se encontró el CIF, no es nuevo
          MsgBox "Ese " & LblOb(2).Caption & " ya existe.", vbInformation
        End If
      End If
    
    Case 3
      If TxtOb(Index) <> "" Then
        If Not IsDate(TxtOb(Index)) Then
          MsgBox "Esa fecha no es válida.", vbInformation
          Cancel = True
        Else
          TxtOb(Index) = Format(TxtOb(Index), "dd/mm/yyyy")
        End If
      End If
    
  End Select
End Sub

'Selecciona el texto de los campos TxtPres al seleccionar su TextBox
'Select all text after entering in the field
Private Sub TxtPres_GotFocus(Index As Integer)
  SeleccionaTexto TxtPres(Index)
End Sub

'Acciones a realizar cuando se presiona una tecla en un TextBox TxtPres
'Checks while pressing a key in the field
Private Sub TxtPres_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case Index
    Case 2, 9
      'Valida las entradas numéricas
      CompruebaEntero KeyAscii

    Case 3
      'Valida las entradas fecha
      CompruebaFecha KeyAscii

  End Select
End Sub

'Validación de un TextBox antes de salir
'Checks while leaving the field
Private Sub TxtPres_Validate(Index As Integer, Cancel As Boolean)
  Select Case Index
    Case 3
      If TxtPres(Index) <> "" Then
        If Not IsDate(TxtPres(Index)) Then
          MsgBox "Esa fecha no es válida.", vbInformation
          Cancel = True
        Else
          TxtPres(Index) = Format(TxtPres(Index), "dd/mm/yyyy")
        End If
      End If

    Case 7
      'Es un email, pasar a minúsculas
      If TxtPres(Index) <> "" Then
        TxtPres(Index) = LCase(TxtPres(Index))
        'Ver si ya existe ese email
        If Busca(Conexion, TxtPres(Index), TxtPres(Index).Tag, Tabla) And blnInsertando Then
          'Se encontró, no es nuevo
          MsgBox "Ese e-mail ya existe.", vbInformation
        End If
      End If
    
  End Select
End Sub

'Activar la fecha de baja
'Activate a hidden field
Private Sub ChkBaja_Click()
  If ChkBaja.Value Then
    LblPres(6).Visible = True
    TxtPres(3).Visible = True
  Else
    LblPres(6).Visible = False
    TxtPres(3).Visible = False
  End If
End Sub

'Cuando se hace dblclick sobre la label de una combo, vaciar ésta
'Clear a FixedCombo while double clicking on it
Private Sub Lbls_DblClick(Index As Integer)
  If blnModificando Or blnInsertando Then
    Select Case Index
      Case 0
        CmbLocalidad.ListIndex = -1
      Case 1
        CmbProvincia.ListIndex = -1
    End Select
  End If
End Sub
