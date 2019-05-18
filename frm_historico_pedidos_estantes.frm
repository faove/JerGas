VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_historico_pedidos_estantes 
   Caption         =   "SISTEMA DE GESTIÓN Y CONTROL JER-GAS, C.A."
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7590
   Icon            =   "frm_historico_pedidos_estantes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      DataField       =   "codigo"
      DataSource      =   "maestro"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc maestro 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
      RecordSource    =   "tbl_estantes"
      Caption         =   "maestro"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "INGRESE  DESCRIPCIÓN  DEL CLIENTE"
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
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   5295
      Begin MSDataListLib.DataCombo Dcmb_Buscar 
         Bindings        =   "frm_historico_pedidos_estantes.frx":08CA
         DataField       =   "codigo"
         DataSource      =   "maestro"
         Height          =   420
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Pulse doble click para cambiar el tipo de busqueda, después de presionar búsqueda avanzada"
         Top             =   720
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   741
         _Version        =   393216
         MatchEntry      =   -1  'True
         ListField       =   "cliente"
         BoundColumn     =   "codigo"
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LABEL_BUSCA 
         BackStyle       =   0  'Transparent
         Caption         =   "Búsqueda:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CERRAR"
      Height          =   735
      Left            =   5760
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROCESAR"
      Height          =   735
      Left            =   5760
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label_titulo 
      BackColor       =   &H80000001&
      Caption         =   "  HISTÓRICO DE PEDIDOS (ESTANTES)"
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
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   8655
   End
End
Attribute VB_Name = "frm_historico_pedidos_estantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
rpt_historico_pedidos_estantes.Show
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
