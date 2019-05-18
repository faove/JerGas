VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_resumen_pedidos 
   Caption         =   "SISTEMA DE GESTIÓN Y CONTROL JER-GAS, C.A."
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5835
   Icon            =   "frm_resumen_pedidos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      DataField       =   "status"
      DataSource      =   "pedidos"
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Text            =   "VI"
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame13 
      Caption         =   "Nº Ruta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   735
      Left            =   2760
      TabIndex        =   5
      ToolTipText     =   "Suministre el Tipo de Contrato del Cliente."
      Top             =   1200
      Width           =   1095
      Begin VB.ComboBox txt_ruta 
         DataField       =   "id_ruta"
         DataSource      =   "rutas"
         Height          =   315
         ItemData        =   "frm_resumen_pedidos.frx":08CA
         Left            =   120
         List            =   "frm_resumen_pedidos.frx":08DA
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "INGRESE FECHA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   975
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   2175
      Begin MSComCtl2.DTPicker DTPicker_desde 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   59244545
         CurrentDate     =   39344
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CERRAR"
      Height          =   735
      Left            =   3000
      TabIndex        =   1
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROCESAR"
      Height          =   735
      Left            =   1440
      TabIndex        =   0
      Top             =   2160
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc rutas 
      Height          =   450
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   794
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=gergas"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "gergas"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tbl_ruta"
      Caption         =   "rutas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc pedidos 
      Height          =   450
      Left            =   1800
      Top             =   0
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   794
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=gergas"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "gergas"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tbl_pedidos"
      Caption         =   "pedidos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label_titulo 
      BackColor       =   &H80000001&
      Caption         =   "  RESUMEN DE PEDIDOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   8655
   End
End
Attribute VB_Name = "frm_resumen_pedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
rpt_resumen_pedidos.Show
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
