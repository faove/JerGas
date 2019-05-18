VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_instalaciones 
   Caption         =   "SISTEMA DE GESTIÓN Y CONTROL JER-GAS, C.A."
   ClientHeight    =   9825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12705
   Icon            =   "frm_instalaciones.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text51 
      Height          =   375
      Left            =   13320
      TabIndex        =   168
      Top             =   5880
      Width           =   1575
   End
   Begin VB.TextBox Text50 
      Height          =   375
      Left            =   13320
      TabIndex        =   167
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox Text49 
      Height          =   375
      Left            =   13320
      TabIndex        =   166
      Top             =   5160
      Width           =   1575
   End
   Begin VB.TextBox Text48 
      Height          =   375
      Left            =   13320
      TabIndex        =   165
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CheckBox Check4 
      Caption         =   "INGRESAR OTROS MATERIALES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   480
      TabIndex        =   164
      Top             =   6960
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Frame Frame_final 
      Caption         =   "TOTAL GENERAL"
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
      Height          =   1125
      Left            =   9600
      TabIndex        =   159
      ToolTipText     =   "Total de Factura "
      Top             =   7560
      Visible         =   0   'False
      Width           =   3135
      Begin VB.Frame Frame32 
         Caption         =   "Total a Pagar"
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
         Height          =   650
         Left            =   1320
         TabIndex        =   162
         ToolTipText     =   "Precio Total"
         Top             =   360
         Width           =   1695
         Begin VB.TextBox txt_monto_general 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   163
            TabStop         =   0   'False
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame31 
         Caption         =   "IVA"
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
         Height          =   650
         Left            =   120
         TabIndex        =   160
         ToolTipText     =   "I.V.A."
         Top             =   360
         Width           =   1095
         Begin VB.TextBox txt_iva_general 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   161
            TabStop         =   0   'False
            Top             =   240
            Width           =   855
         End
      End
   End
   Begin VB.CheckBox Check3 
      Caption         =   "COSTO ADICIONAL DE MATERIALES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   158
      Top             =   6960
      Width           =   3855
   End
   Begin VB.Frame Frame_adicional 
      Caption         =   "INGRESAR MATERIAL ADICIONAL"
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
      Height          =   1125
      Left            =   240
      TabIndex        =   147
      Top             =   7560
      Visible         =   0   'False
      Width           =   9255
      Begin VB.Frame Frame28 
         Caption         =   "Cant."
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
         Height          =   650
         Left            =   3480
         TabIndex        =   156
         ToolTipText     =   "Cantidad "
         Top             =   360
         Width           =   855
         Begin VB.TextBox txt_cant_a 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   157
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame27 
         Caption         =   "Total a Pagar"
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
         Height          =   650
         Left            =   7440
         TabIndex        =   154
         ToolTipText     =   "Total a Pagar"
         Top             =   360
         Width           =   1695
         Begin VB.TextBox txt_total_a 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   155
            TabStop         =   0   'False
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame26 
         Caption         =   "IVA"
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
         Height          =   650
         Left            =   6240
         TabIndex        =   152
         ToolTipText     =   "I.V.A."
         Top             =   360
         Width           =   1095
         Begin VB.TextBox txt_iva_a 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   153
            TabStop         =   0   'False
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame25 
         Caption         =   "Precio Unit."
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
         Height          =   650
         Left            =   4440
         TabIndex        =   150
         ToolTipText     =   "Precio Unitario"
         Top             =   360
         Width           =   1695
         Begin VB.TextBox txt_precio_a 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   151
            TabStop         =   0   'False
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame24 
         Caption         =   "DESCRIPCIÓN"
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
         Height          =   650
         Left            =   120
         TabIndex        =   148
         ToolTipText     =   "Descripción"
         Top             =   360
         Width           =   3255
         Begin MSDataListLib.DataCombo txt_descripcion 
            Bindings        =   "frm_instalaciones.frx":08CA
            Height          =   315
            Left            =   120
            TabIndex        =   149
            Top             =   240
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "descripcion"
            BoundColumn     =   "codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin VB.TextBox Text47 
      DataField       =   "id_pedido"
      DataSource      =   "nuevainstalaciones"
      Height          =   285
      Left            =   6480
      TabIndex        =   146
      Text            =   "Text47"
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame22 
      Caption         =   "Control de Factura"
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
      Height          =   700
      Left            =   9840
      TabIndex        =   143
      ToolTipText     =   "Suministre la Cédula del Cliente."
      Top             =   2040
      Width           =   3735
      Begin VB.TextBox txt_factura 
         Alignment       =   2  'Center
         DataField       =   "id_factura"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "control_procesos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   145
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txt_control 
         Alignment       =   2  'Center
         DataField       =   "id_control"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "control_procesos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   144
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox txt_mat8 
      DataField       =   "rabo"
      DataSource      =   "inventario"
      Height          =   405
      Left            =   13680
      TabIndex        =   138
      Top             =   9600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame_descripcion 
      Caption         =   "DESCRIPCIÓN"
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
      Height          =   1125
      Left            =   240
      TabIndex        =   127
      Top             =   5640
      Width           =   12495
      Begin VB.Frame Frame21 
         Caption         =   "Precio GLP"
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
         Height          =   650
         Left            =   6360
         TabIndex        =   141
         ToolTipText     =   "Precio GLP"
         Top             =   360
         Width           =   1215
         Begin VB.TextBox txt_total_cilindro 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   142
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Sub Total"
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
         Height          =   650
         Left            =   7680
         TabIndex        =   139
         ToolTipText     =   "Sub-Total"
         Top             =   360
         Width           =   1695
         Begin VB.TextBox txt_monto_fac 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   140
            TabStop         =   0   'False
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Tipo de Instalación"
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
         Height          =   650
         Left            =   120
         TabIndex        =   136
         ToolTipText     =   "Descripción"
         Top             =   360
         Width           =   3375
         Begin VB.TextBox txt_instalacion 
            Alignment       =   2  'Center
            DataField       =   "descripcion"
            DataSource      =   "facturando"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   137
            TabStop         =   0   'False
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "Precio de Inst"
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
         Height          =   650
         Left            =   4560
         TabIndex        =   134
         ToolTipText     =   "Precio de Instalación"
         Top             =   360
         Width           =   1695
         Begin VB.TextBox txt_precio_uni 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   135
            TabStop         =   0   'False
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "IVA"
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
         Height          =   650
         Left            =   9480
         TabIndex        =   132
         ToolTipText     =   "I.V.A."
         Top             =   360
         Width           =   1095
         Begin VB.TextBox txt_iva 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   133
            TabStop         =   0   'False
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Total a Pagar"
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
         Height          =   650
         Left            =   10680
         TabIndex        =   130
         ToolTipText     =   "Total a Pagar"
         Top             =   360
         Width           =   1695
         Begin VB.TextBox txt_total_pagar 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   131
            TabStop         =   0   'False
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Cant."
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
         Height          =   650
         Left            =   3600
         TabIndex        =   128
         ToolTipText     =   "Cantidad de Cilindros"
         Top             =   360
         Width           =   855
         Begin VB.TextBox txt_cant 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   129
            Top             =   240
            Width           =   615
         End
      End
   End
   Begin VB.TextBox Text43 
      DataField       =   "id_pedido"
      DataSource      =   "facturando"
      Height          =   285
      Left            =   9600
      TabIndex        =   126
      Text            =   "Text43"
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txt_actual8 
      Height          =   375
      Left            =   14280
      TabIndex        =   124
      TabStop         =   0   'False
      Top             =   9600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text41 
      Height          =   375
      Left            =   11640
      TabIndex        =   118
      Text            =   "Text41"
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text33 
      DataField       =   "cant_cilindro"
      DataSource      =   "NUEVASINST"
      Height          =   285
      Left            =   14040
      TabIndex        =   110
      Text            =   "Text33"
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txt_venta 
      DataField       =   "precio_cilindro"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   " #.##0,00;( #.##0,00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   8202
         SubFormatType   =   0
      EndProperty
      DataSource      =   "instalacion"
      Height          =   285
      Left            =   10440
      TabIndex        =   109
      Text            =   "Text33"
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt_fecha_pedido 
      Height          =   285
      Left            =   13560
      TabIndex        =   108
      Text            =   "Text33"
      Top             =   10440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txt_pedidos 
      DataField       =   "id_pedido"
      DataSource      =   "pedidos"
      Height          =   285
      Left            =   13680
      TabIndex        =   107
      Top             =   10080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text31 
      DataField       =   "codigo"
      DataSource      =   "NUEVASINST"
      Height          =   285
      Left            =   14040
      TabIndex        =   104
      Text            =   "Text31"
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSAdodcLib.Adodc nuevainstalaciones 
      Height          =   330
      Left            =   2160
      Top             =   360
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
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
      RecordSource    =   "tbl_resumen_instalaciones"
      Caption         =   "nuevainstalaciones"
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
   Begin VB.TextBox Text30 
      DataField       =   "codigo"
      DataSource      =   "resumen_inv2"
      Height          =   285
      Left            =   14040
      TabIndex        =   94
      Text            =   "Text30"
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text29 
      DataField       =   "codigo"
      DataSource      =   "resumen_inv1"
      Height          =   285
      Left            =   14040
      TabIndex        =   93
      Text            =   "Text29"
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text13 
      DataField       =   "codigo"
      DataSource      =   "materiales"
      Height          =   375
      Left            =   11880
      TabIndex        =   92
      Text            =   "Text13"
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSAdodcLib.Adodc materiales 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
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
      RecordSource    =   "tbl_materiales"
      Caption         =   "materiales"
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
   Begin VB.TextBox txt_actual7 
      Height          =   375
      Left            =   14280
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   9240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt_actual4 
      Height          =   375
      Left            =   14280
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   8160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt_actual5 
      Height          =   375
      Left            =   14280
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   8520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt_actual6 
      Height          =   375
      Left            =   14280
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   8880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt_actual1 
      Height          =   375
      Left            =   14280
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   7080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt_actual2 
      Height          =   375
      Left            =   14280
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   7440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt_actual3 
      Height          =   375
      Left            =   14280
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt_mat7 
      DataField       =   "tubo"
      DataSource      =   "inventario"
      Height          =   375
      Left            =   13680
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   9240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt_mat4 
      DataField       =   "copa"
      DataSource      =   "inventario"
      Height          =   375
      Left            =   13680
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   8160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt_mat5 
      DataField       =   "reductor"
      DataSource      =   "inventario"
      Height          =   375
      Left            =   13680
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   8520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt_mat6 
      DataField       =   "tee"
      DataSource      =   "inventario"
      Height          =   375
      Left            =   13680
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   8880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt_mat3 
      DataField       =   "conector"
      DataSource      =   "inventario"
      Height          =   375
      Left            =   13680
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt_mat2 
      DataField       =   "regulador"
      DataSource      =   "inventario"
      Height          =   375
      Left            =   13680
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   7440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt_mat1 
      DataField       =   "cilindro"
      DataSource      =   "inventario"
      Height          =   375
      Left            =   13680
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   7080
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSAdodcLib.Adodc kit 
      Height          =   375
      Left            =   5760
      Top             =   1320
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      RecordSource    =   "tbl_kit"
      Caption         =   "kit"
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
   Begin VB.TextBox Text12 
      DataField       =   "fecha_ult_pago"
      DataSource      =   "clientes"
      Height          =   285
      Left            =   12480
      TabIndex        =   56
      Text            =   "Text12"
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Asignar Kit de Instalación"
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
      Left            =   480
      TabIndex        =   11
      Top             =   8760
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSAdodcLib.Adodc inventario 
      Height          =   330
      Left            =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
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
      RecordSource    =   "tbl_inventario"
      Caption         =   "inventario"
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
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   480
   End
   Begin VB.TextBox Text4 
      DataField       =   "fecha_ini_contrato"
      DataSource      =   "Clientes"
      Height          =   375
      Left            =   12480
      TabIndex        =   47
      TabStop         =   0   'False
      Text            =   "Text4"
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   12480
      TabIndex        =   46
      TabStop         =   0   'False
      Text            =   "Text2"
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text3 
      DataField       =   "id"
      DataSource      =   "control_clientes"
      Height          =   375
      Left            =   12480
      TabIndex        =   45
      TabStop         =   0   'False
      Text            =   "Text3"
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSAdodcLib.Adodc clientes 
      Height          =   330
      Left            =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
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
      RecordSource    =   "tbl_clientes"
      Caption         =   "clientes"
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
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   12480
      TabIndex        =   44
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame17 
      Height          =   975
      Left            =   6840
      TabIndex        =   41
      Top             =   4560
      Width           =   5895
      Begin VB.CommandButton cmdeliminar 
         Caption         =   "&Eliminar"
         Height          =   615
         Left            =   3000
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdagregar 
         Caption         =   "&Agregar"
         Height          =   615
         Left            =   120
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Botón para Agregar un Nuevo Usuario"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdcancelar 
         Caption         =   "&Cancelar"
         Height          =   615
         Left            =   120
         TabIndex        =   42
         ToolTipText     =   "Pulse este botón si desea Cancelar el Usuario Agregado"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdguardar 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   615
         Left            =   1560
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Para Salvar el Usuario Agregado o Modificado"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdsalir 
         Caption         =   "C&errar"
         Height          =   615
         Left            =   4440
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Cerrar el Sistema"
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Código de Cliente"
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
      Height          =   700
      Left            =   240
      TabIndex        =   39
      ToolTipText     =   "Suministre el Código del Cliente"
      Top             =   2040
      Width           =   1815
      Begin VB.TextBox txt_codigo 
         Alignment       =   2  'Center
         DataField       =   "codigo"
         DataSource      =   "clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Fecha de Contrato"
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
      Height          =   700
      Left            =   3960
      TabIndex        =   38
      ToolTipText     =   "Suministre la Fecha del Contrato del Cliente."
      Top             =   2040
      Width           =   2055
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "fecha_ini_contrato"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   3
         EndProperty
         DataSource      =   "clientes"
         Height          =   345
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   66519041
         CurrentDate     =   39083
         MinDate         =   39083
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Teléfono Habitación"
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
      Height          =   700
      Left            =   9840
      TabIndex        =   37
      ToolTipText     =   "Suministre los Teléfonos del Cliente."
      Top             =   2880
      Width           =   2415
      Begin VB.TextBox txt_telefono_hab 
         Alignment       =   2  'Center
         DataField       =   "telefono_hab"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "(####) ### ## ##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame33 
      Caption         =   "Apellidos y Nombres"
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
      Height          =   700
      Left            =   240
      TabIndex        =   36
      ToolTipText     =   "Suministre Apellido y Nombre del Cliente"
      Top             =   2880
      Width           =   7215
      Begin VB.TextBox txt_cliente 
         DataField       =   "cliente"
         DataSource      =   "clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   6975
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Nº de Contrato"
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
      Height          =   700
      Left            =   2280
      TabIndex        =   34
      ToolTipText     =   "Suministre el Nº de Contrato del Cliente"
      Top             =   2040
      Width           =   1575
      Begin VB.TextBox txt_contrato 
         Alignment       =   2  'Center
         DataField       =   "contrato_num"
         DataSource      =   "clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Cédula de Identidad"
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
      Height          =   700
      Left            =   7560
      TabIndex        =   33
      ToolTipText     =   "Suministre la Cédula del Cliente."
      Top             =   2880
      Width           =   2055
      Begin VB.TextBox txt_cedula 
         Alignment       =   2  'Center
         DataField       =   "cedula"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Dirección"
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
      Height          =   700
      Left            =   240
      TabIndex        =   32
      ToolTipText     =   "Suministre la Dirección del Cliente."
      Top             =   3720
      Width           =   9375
      Begin VB.TextBox txt_direccion 
         DataField       =   "direccion"
         DataSource      =   "clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   9135
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Correo Electrónico"
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
      Height          =   700
      Left            =   9840
      TabIndex        =   31
      ToolTipText     =   "Suministre el Correo Electrónico del Cliente."
      Top             =   3720
      Width           =   5175
      Begin VB.TextBox txt_correo 
         DataField       =   "e_mail"
         DataSource      =   "clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Observaciones"
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
      Height          =   855
      Left            =   240
      TabIndex        =   30
      ToolTipText     =   "Suministre Cualquier Dato Adicional del Cliente."
      Top             =   4560
      Width           =   5415
      Begin VB.TextBox txt_observaciones 
         DataField       =   "observaciones"
         DataSource      =   "clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Frame Frame14 
      Caption         =   "Teléfono Celular"
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
      Height          =   700
      Left            =   12600
      TabIndex        =   29
      ToolTipText     =   "Suministre los Teléfonos del Cliente."
      Top             =   2880
      Width           =   2415
      Begin VB.TextBox txt_telefono_cel 
         Alignment       =   2  'Center
         DataField       =   "telefono_cel"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "(####) ### ## ##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "T/Contr."
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
      Height          =   700
      Left            =   7320
      TabIndex        =   28
      ToolTipText     =   "Suministre el Tipo de Contrato del Cliente."
      Top             =   2040
      Width           =   1095
      Begin MSDataListLib.DataCombo txt_cilindro 
         Bindings        =   "frm_instalaciones.frx":08E3
         DataField       =   "id_inst"
         DataSource      =   "clientes"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         ListField       =   "id_inst"
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Status"
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
      Height          =   700
      Left            =   6120
      TabIndex        =   27
      ToolTipText     =   "Suministre el Status del Cliente."
      Top             =   2040
      Width           =   1095
      Begin MSDataListLib.DataCombo txt_status 
         Bindings        =   "frm_instalaciones.frx":08FD
         DataField       =   "status"
         DataSource      =   "clientes"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         ListField       =   "status"
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "Rutas"
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
      Height          =   700
      Left            =   8520
      TabIndex        =   26
      ToolTipText     =   "Suministre el Tipo de Contrato del Cliente."
      Top             =   2040
      Width           =   1095
      Begin MSDataListLib.DataCombo txt_ruta 
         Bindings        =   "frm_instalaciones.frx":0912
         DataField       =   "id_ruta"
         DataSource      =   "clientes"
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         ListField       =   "id_ruta"
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc status 
      Height          =   330
      Left            =   5760
      Top             =   840
      Visible         =   0   'False
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   582
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
      RecordSource    =   "tbl_status"
      Caption         =   "status"
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
   Begin MSAdodcLib.Adodc rutas 
      Height          =   330
      Left            =   2160
      Top             =   1440
      Visible         =   0   'False
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   582
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
   Begin MSAdodcLib.Adodc control_clientes 
      Height          =   330
      Left            =   2160
      Top             =   720
      Visible         =   0   'False
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   582
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
      RecordSource    =   "tbl_control_clientes"
      Caption         =   "control_clientes"
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
   Begin MSAdodcLib.Adodc instalacion 
      Height          =   330
      Left            =   2160
      Top             =   1080
      Visible         =   0   'False
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   582
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
      RecordSource    =   "tbl_instalacion"
      Caption         =   "instalacion"
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
   Begin MSAdodcLib.Adodc resumen_inv1 
      Height          =   330
      Left            =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
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
      RecordSource    =   "tbl_resumen_inventario"
      Caption         =   "resumen_inv1"
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
   Begin MSAdodcLib.Adodc resumen_inv2 
      Height          =   330
      Left            =   -120
      Top             =   1080
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
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
      RecordSource    =   "tbl_resumen_inventario2"
      Caption         =   "resumen_inv2"
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
   Begin MSAdodcLib.Adodc pedidos 
      Height          =   330
      Left            =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame_kit 
      Caption         =   "Kit de Instalación"
      Enabled         =   0   'False
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
      Height          =   3615
      Left            =   3600
      TabIndex        =   48
      Top             =   960
      Visible         =   0   'False
      Width           =   12495
      Begin VB.TextBox Text46 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text45 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "tubo"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "kit"
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   122
         Text            =   "Text45"
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text44 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "tot_rabo"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "NUEVASINST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   121
         Text            =   "Text44"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txt_iva8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "rabo"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "kit"
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   120
         Text            =   "IVA8"
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text42 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         DataField       =   "cant_rabo"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "NUEVASINST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   119
         Top             =   1200
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text40 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         DataField       =   "cant_cilindro"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "NUEVASINST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   117
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text39 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         DataField       =   "cant_regulador"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "NUEVASINST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   116
         Top             =   840
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text38 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         DataField       =   "cant_conector"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "NUEVASINST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   115
         Top             =   1560
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text37 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         DataField       =   "cant_copa"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "NUEVASINST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   114
         Top             =   1920
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text36 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         DataField       =   "cant_reductor"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "NUEVASINST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   113
         Top             =   2280
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text35 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         DataField       =   "cant_tee"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "NUEVASINST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   112
         Top             =   2640
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text34 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         DataField       =   "cant_tubo"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "NUEVASINST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   111
         Top             =   3000
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txt_iva1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "cilindro"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "kit"
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   103
         Text            =   "IVA1"
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txt_iva2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "regulador"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "kit"
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   102
         Text            =   "IVA2"
         Top             =   960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txt_iva3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "conector"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "kit"
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   101
         Text            =   "IVA3"
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txt_iva4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "copa"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "kit"
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   100
         Text            =   "IVA4"
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txt_iva5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "reductor"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "kit"
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   99
         Text            =   "IVA5"
         Top             =   2400
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txt_iva6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "tee"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "kit"
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   98
         Text            =   "IVA6"
         Top             =   2760
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txt_iva7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "tubo"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "kit"
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   97
         Text            =   "IVA7"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame Frame1 
         Caption         =   "Cambiar Parámetros"
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
         Height          =   1095
         Left            =   9000
         TabIndex        =   95
         Top             =   1920
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton Command1 
            Caption         =   "Aceptar"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Ingresar Cambios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Visible         =   0   'False
            Width           =   1815
         End
      End
      Begin VB.TextBox Text22 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "tot_cilindro"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "NUEVASINST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   80
         Text            =   "Text22"
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmd_imprime 
         Caption         =   "Imprimir Contrato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11040
         TabIndex        =   75
         Top             =   1080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Frame Frame_total 
         Caption         =   "Total a Cancelar"
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
         Left            =   9000
         TabIndex        =   73
         Top             =   600
         Visible         =   0   'False
         Width           =   2055
         Begin VB.TextBox Text32 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            TabIndex        =   105
            Top             =   300
            Width           =   1215
         End
         Begin VB.TextBox Text21 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   74
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Caption         =   "IVA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   106
            Top             =   400
            Width           =   375
         End
      End
      Begin VB.TextBox Text28 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "tot_tubo"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "NUEVASINST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   71
         Text            =   "Text28"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text27 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "tot_tee"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "NUEVASINST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   70
         Text            =   "Text27"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text26 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "tot_reductor"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "NUEVASINST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   69
         Text            =   "Text26"
         Top             =   2400
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text25 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "tot_copa"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "NUEVASINST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   68
         Text            =   "Text25"
         Top             =   2040
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text24 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "tot_conector"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "NUEVASINST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   67
         Text            =   "Text24"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text23 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "tot_regulador"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "NUEVASINST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   66
         Text            =   "Text23"
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "tubo"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "kit"
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   63
         Text            =   "Text20"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text19 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "reductor"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "kit"
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   62
         Text            =   "Text19"
         Top             =   2760
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text18 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "copa"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "kit"
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   61
         Text            =   "Text18"
         Top             =   2400
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "conector"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "kit"
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   60
         Text            =   "Text17"
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "tee"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "kit"
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   59
         Text            =   "Text16"
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "regulador"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "kit"
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   58
         Text            =   "Text15"
         Top             =   960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text14 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DataField       =   "cilindro"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "kit"
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "Text14"
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   3120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   2760
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   2040
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label15 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   125
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label14 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   123
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "IVA: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   96
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
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
         Height          =   255
         Left            =   8640
         TabIndex        =   76
         Top             =   240
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label10 
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   72
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   65
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Precio Unitario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   64
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   55
         Top             =   3120
         Width           =   3615
      End
      Begin VB.Label Label6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   54
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label Label5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   53
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Label Label4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   52
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label Label3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   51
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   50
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   49
         Top             =   600
         Width           =   2535
      End
   End
   Begin MSAdodcLib.Adodc facturando 
      Height          =   330
      Left            =   2160
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
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
      RecordSource    =   "tbl_facturando"
      Caption         =   "facturando"
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
   Begin MSAdodcLib.Adodc control_procesos 
      Height          =   330
      Left            =   4200
      Top             =   0
      Visible         =   0   'False
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   582
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
      RecordSource    =   "tbl_control_procesos"
      Caption         =   "control_clientes"
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
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   960
      Width           =   15465
   End
   Begin VB.Label Label_titulo 
      BackColor       =   &H80000001&
      Caption         =   "  CONTROL DE NUEVAS INSTALACIONES"
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
      Left            =   7560
      TabIndex        =   43
      Top             =   240
      Width           =   7695
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000B&
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   0
      Top             =   900
      Width           =   15345
   End
End
Attribute VB_Name = "frm_instalaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cant_venta, cant_ventaA
Dim acumula1, acumula2, acumula3
Dim paso1, paso2, paso3, paso4, paso5, paso6, pasoFinal, pasoprevio
Dim paso1A, paso2A, paso3A, paso4A, paso5A, paso6A, pasoFinalA, pasoprevioA

Private Sub Check3_Click()
  If Check3.Value = 1 Then
     Frame_adicional.Visible = True
  End If
  
  If Check3.Value = 0 Then
     Frame_adicional.Visible = False
     Me.txt_descripcion.Text = ""
     Me.txt_cant_a.Text = ""
     Me.txt_precio_a.Text = ""
     Me.txt_iva_a.Text = ""
     Me.txt_total_a.Text = ""
  End If
End Sub

Private Sub Check4_Click()
  
  If Check4.Value = 1 Then
      Me.txt_descripcion.Text = ""
      Me.txt_cant_a.Text = ""
      Me.txt_precio_a.Text = ""
      Me.txt_iva_a.Text = ""
      Me.txt_total_a.Text = ""
  End If
End Sub

Private Sub cmdguardar_Click()
Dim monto
Dim concepto
Dim monto_cilindro
Dim total_iva
Dim pedido As Integer
 
 If txt_cilindro = "10" Then
     
     concepto = "Suministro de GLP (10 Kgs)"
     
     Me.inventario.CommandType = adCmdText
 Me.inventario.RecordSource = "select * from tbl_inventario WHERE tbl_inventario.id_inst = 10 ORDER BY id_inst ASC"

 Me.txt_actual1.Text = (Me.txt_mat1.Text - Me.Text5.Text)
 Me.txt_actual2.Text = (Me.txt_mat2.Text - Me.Text6.Text)
 Me.txt_actual3.Text = (Me.txt_mat3.Text - Me.Text7.Text)
 Me.txt_actual4.Text = (Me.txt_mat4.Text - Me.Text8.Text)
 Me.txt_actual5.Text = (Me.txt_mat5.Text - Me.Text9.Text)
 Me.txt_actual6.Text = (Me.txt_mat6.Text - Me.Text10.Text)
 Me.txt_actual7.Text = (Me.txt_mat7.Text - Me.Text11.Text)
 Me.txt_actual8.Text = (Me.txt_mat8.Text - Me.Text46.Text)
 
      inventario.Recordset!cilindro = Me.txt_actual1.Text
      inventario.Recordset!regulador = Me.txt_actual2.Text
      inventario.Recordset!conector = Me.txt_actual3.Text
      inventario.Recordset!copa = Me.txt_actual4.Text
      inventario.Recordset!reductor = Me.txt_actual5.Text
      inventario.Recordset!tee = Me.txt_actual6.Text
      inventario.Recordset!tubo = Me.txt_actual7.Text
      inventario.Recordset!rabo = Me.txt_actual8.Text
      inventario.Recordset.Update
 Me.inventario.Refresh
     
     With materiales.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
            .Find "codigo = 1"
               !cant_actual = CInt(Me.txt_actual1.Text)
            .Find "codigo = 5"
               !cant_actual = CInt(Me.txt_actual2.Text)
            .Find "codigo = 9"
               !cant_actual = !cant_actual - CInt(Me.Text46.Text)
            .Find "codigo = 10"
               !cant_actual = !cant_actual - CInt(Me.Text7.Text)
            .Find "codigo = 11"
               !cant_actual = !cant_actual - CInt(Me.Text8.Text)
            .Find "codigo = 12"
               !cant_actual = !cant_actual - CInt(Me.Text9.Text)
            .Find "codigo = 13"
               !cant_actual = !cant_actual - CInt(Me.Text10.Text)
            .Find "codigo = 14"
               !cant_actual = !cant_actual - CInt(Me.Text11.Text)
            .Update
      End With
      
      With resumen_inv1.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "codigo = 1"
              !cil_lleno = !cil_lleno - CInt(Me.Text5.Text)
              !cil_vacio = !cil_vacio + CInt(Me.Text5.Text)
        .Update
      End With
        Me.resumen_inv1.Refresh
      
      With resumen_inv2.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
            .Find "codigo = 5"
               !cant_actual = CInt(Me.txt_actual2.Text)
               !cant_inst = !cant_inst + CInt(Me.Text6.Text)
            .Find "codigo = 9"
               !cant_actual = !cant_actual - CInt(Me.Text46.Text)
               !cant_inst = !cant_inst + CInt(Me.Text46.Text)
            .Find "codigo = 10"
               !cant_actual = !cant_actual - CInt(Me.Text7.Text)
               !cant_inst = !cant_inst + CInt(Me.Text7.Text)
            .Find "codigo = 11"
               !cant_actual = !cant_actual - CInt(Me.Text8.Text)
               !cant_inst = !cant_inst + CInt(Me.Text8.Text)
            .Find "codigo = 12"
               !cant_actual = !cant_actual - CInt(Me.Text9.Text)
               !cant_inst = !cant_inst + CInt(Me.Text9.Text)
            .Find "codigo = 13"
               !cant_actual = !cant_actual - CInt(Me.Text10.Text)
               !cant_inst = !cant_inst + CInt(Me.Text10.Text)
            .Find "codigo = 14"
               !cant_actual = !cant_actual - CInt(Me.Text11.Text)
               !cant_inst = !cant_inst + CInt(Me.Text11.Text)
            .Update
      End With
  
 
 Call actualizar_cn("SQL Server")
     
     Gcod_planilla = FGNRO_LIQ()
     Gcod_control = FGNRO_CONTROL()
     Gcod_factura = FGNRO_FACTURA()
        
 With pedidos.Recordset
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
        .MoveLast
    End If
        .AddNew
    !usuario_liq = Usuario
    !id_pedido = Gcod_planilla
    !num_control = Gcod_control
    !num_factura = Gcod_factura
    !codigo = txt_codigo.Text
    !status = "VI"
    !fecha_pedido = Me.txt_fecha_pedido.Text
    !id_inst = Me.txt_cilindro.Text
    !monto_fac = CCur(Me.txt_total_cilindro.Text)
    !cant_pedido = CInt(Me.txt_cant.Text)
       .Update
                                                                                
  With clientes.Recordset
       !fecha_ult_pago = Me.txt_fecha_pedido.Text
          .Update
   End With
  End With
End If

If txt_cilindro = "18" Then
     
 concepto = "Suministro de GLP (18 Kgs)"
 
 Me.inventario.CommandType = adCmdText
 Me.inventario.RecordSource = "select * from tbl_inventario WHERE tbl_inventario.id_inst = 18 ORDER BY id_inst ASC"

 Me.txt_actual1.Text = (Me.txt_mat1.Text - Me.Text5.Text)
 Me.txt_actual2.Text = (Me.txt_mat2.Text - Me.Text6.Text)
 Me.txt_actual3.Text = (Me.txt_mat3.Text - Me.Text7.Text)
 Me.txt_actual4.Text = (Me.txt_mat4.Text - Me.Text8.Text)
 Me.txt_actual5.Text = (Me.txt_mat5.Text - Me.Text9.Text)
 Me.txt_actual6.Text = (Me.txt_mat6.Text - Me.Text10.Text)
 Me.txt_actual7.Text = (Me.txt_mat7.Text - Me.Text11.Text)
 Me.txt_actual8.Text = (Me.txt_mat8.Text - Me.Text46.Text)
 
      inventario.Recordset!cilindro = Me.txt_actual1.Text
      inventario.Recordset!regulador = Me.txt_actual2.Text
      inventario.Recordset!conector = Me.txt_actual3.Text
      inventario.Recordset!copa = Me.txt_actual4.Text
      inventario.Recordset!reductor = Me.txt_actual5.Text
      inventario.Recordset!tee = Me.txt_actual6.Text
      inventario.Recordset!tubo = Me.txt_actual7.Text
      inventario.Recordset!rabo = Me.txt_actual8.Text
      inventario.Recordset.Update
 Me.inventario.Refresh
     
     With materiales.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
            .Find "codigo = 1"
               !cant_actual = CInt(Me.txt_actual1.Text)
            .Find "codigo = 5"
               !cant_actual = CInt(Me.txt_actual2.Text)
            .Find "codigo = 9"
               !cant_actual = !cant_actual - CInt(Me.Text46.Text)
            .Find "codigo = 10"
               !cant_actual = !cant_actual - CInt(Me.Text7.Text)
            .Find "codigo = 11"
               !cant_actual = !cant_actual - CInt(Me.Text8.Text)
            .Find "codigo = 12"
               !cant_actual = !cant_actual - CInt(Me.Text9.Text)
            .Find "codigo = 13"
               !cant_actual = !cant_actual - CInt(Me.Text10.Text)
            .Find "codigo = 14"
               !cant_actual = !cant_actual - CInt(Me.Text11.Text)
            .Update
      End With
      
      With resumen_inv1.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "codigo = 1"
              !cil_lleno = !cil_lleno - CInt(Me.Text5.Text)
              !cil_vacio = !cil_vacio + CInt(Me.Text5.Text)
        .Update
      End With
        Me.resumen_inv1.Refresh
      
      With resumen_inv2.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
            .Find "codigo = 5"
               !cant_actual = CInt(Me.txt_actual2.Text)
               !cant_inst = !cant_inst + CInt(Me.Text6.Text)
            .Find "codigo = 9"
               !cant_actual = !cant_actual - CInt(Me.Text46.Text)
               !cant_inst = !cant_inst + CInt(Me.Text46.Text)
            .Find "codigo = 10"
               !cant_actual = !cant_actual - CInt(Me.Text7.Text)
               !cant_inst = !cant_inst + CInt(Me.Text7.Text)
            .Find "codigo = 11"
               !cant_actual = !cant_actual - CInt(Me.Text8.Text)
               !cant_inst = !cant_inst + CInt(Me.Text8.Text)
            .Find "codigo = 12"
               !cant_actual = !cant_actual - CInt(Me.Text9.Text)
               !cant_inst = !cant_inst + CInt(Me.Text9.Text)
            .Find "codigo = 13"
               !cant_actual = !cant_actual - CInt(Me.Text10.Text)
               !cant_inst = !cant_inst + CInt(Me.Text10.Text)
            .Find "codigo = 14"
               !cant_actual = !cant_actual - CInt(Me.Text11.Text)
               !cant_inst = !cant_inst + CInt(Me.Text11.Text)
            .Update
      End With
  
 
 Call actualizar_cn("SQL Server")
     
     Gcod_planilla = FGNRO_LIQ()
     Gcod_control = FGNRO_CONTROL()
     Gcod_factura = FGNRO_FACTURA()
        
With pedidos.Recordset
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
        .MoveLast
    End If
        .AddNew
    !usuario_liq = Usuario
    !id_pedido = Gcod_planilla
    !num_control = Gcod_control
    !num_factura = Gcod_factura
    
    !codigo = txt_codigo.Text
    !status = "VI"
    !fecha_pedido = Me.txt_fecha_pedido.Text
    !id_inst = Me.txt_cilindro.Text
    !monto_fac = CCur(Me.txt_total_cilindro.Text)
    !cant_pedido = CInt(Me.txt_cant.Text)
       .Update
                                                                                
   With clientes.Recordset
       !fecha_ult_pago = Me.txt_fecha_pedido.Text
          .Update
   End With
  End With
End If
 
 If txt_cilindro = "43" Then
    
 concepto = "Suministro de GLP (43 Kgs)"
 
     Me.inventario.CommandType = adCmdText
 Me.inventario.RecordSource = "select * from tbl_inventario WHERE tbl_inventario.id_inst = 43 ORDER BY id_inst ASC"

 Me.txt_actual1.Text = (Me.txt_mat1.Text - Me.Text5.Text)
 Me.txt_actual2.Text = (Me.txt_mat2.Text - Me.Text6.Text)
 Me.txt_actual3.Text = (Me.txt_mat3.Text - Me.Text7.Text)
 Me.txt_actual4.Text = (Me.txt_mat4.Text - Me.Text8.Text)
 Me.txt_actual5.Text = (Me.txt_mat5.Text - Me.Text9.Text)
 Me.txt_actual6.Text = (Me.txt_mat6.Text - Me.Text10.Text)
 Me.txt_actual7.Text = (Me.txt_mat7.Text - Me.Text11.Text)
 Me.txt_actual8.Text = (Me.txt_mat8.Text - Me.Text46.Text)
 
      inventario.Recordset!cilindro = Me.txt_actual1.Text
      inventario.Recordset!regulador = Me.txt_actual2.Text
      inventario.Recordset!conector = Me.txt_actual3.Text
      inventario.Recordset!copa = Me.txt_actual4.Text
      inventario.Recordset!reductor = Me.txt_actual5.Text
      inventario.Recordset!tee = Me.txt_actual6.Text
      inventario.Recordset!tubo = Me.txt_actual7.Text
      inventario.Recordset!rabo = Me.txt_actual8.Text
      inventario.Recordset.Update
 Me.inventario.Refresh
     
     With materiales.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
            .Find "codigo = 1"
               !cant_actual = CInt(Me.txt_actual1.Text)
            .Find "codigo = 5"
               !cant_actual = CInt(Me.txt_actual2.Text)
            .Find "codigo = 9"
               !cant_actual = !cant_actual - CInt(Me.Text46.Text)
            .Find "codigo = 10"
               !cant_actual = !cant_actual - CInt(Me.Text7.Text)
            .Find "codigo = 11"
               !cant_actual = !cant_actual - CInt(Me.Text8.Text)
            .Find "codigo = 12"
               !cant_actual = !cant_actual - CInt(Me.Text9.Text)
            .Find "codigo = 13"
               !cant_actual = !cant_actual - CInt(Me.Text10.Text)
            .Find "codigo = 14"
               !cant_actual = !cant_actual - CInt(Me.Text11.Text)
            .Update
      End With
      
      With resumen_inv1.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "codigo = 1"
              !cil_lleno = !cil_lleno - CInt(Me.Text5.Text)
              !cil_vacio = !cil_vacio + CInt(Me.Text5.Text)
        .Update
      End With
        Me.resumen_inv1.Refresh
      
      With resumen_inv2.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
            .Find "codigo = 5"
               !cant_actual = CInt(Me.txt_actual2.Text)
               !cant_inst = !cant_inst + CInt(Me.Text6.Text)
            .Find "codigo = 9"
               !cant_actual = !cant_actual - CInt(Me.Text46.Text)
               !cant_inst = !cant_inst + CInt(Me.Text46.Text)
            .Find "codigo = 10"
               !cant_actual = !cant_actual - CInt(Me.Text7.Text)
               !cant_inst = !cant_inst + CInt(Me.Text7.Text)
            .Find "codigo = 11"
               !cant_actual = !cant_actual - CInt(Me.Text8.Text)
               !cant_inst = !cant_inst + CInt(Me.Text8.Text)
            .Find "codigo = 12"
               !cant_actual = !cant_actual - CInt(Me.Text9.Text)
               !cant_inst = !cant_inst + CInt(Me.Text9.Text)
            .Find "codigo = 13"
               !cant_actual = !cant_actual - CInt(Me.Text10.Text)
               !cant_inst = !cant_inst + CInt(Me.Text10.Text)
            .Find "codigo = 14"
               !cant_actual = !cant_actual - CInt(Me.Text11.Text)
               !cant_inst = !cant_inst + CInt(Me.Text11.Text)
            .Update
      End With
  
 
 Call actualizar_cn("SQL Server")
     
     Gcod_planilla = FGNRO_LIQ()
     Gcod_control = FGNRO_CONTROL()
     Gcod_factura = FGNRO_FACTURA()
        
With pedidos.Recordset
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
        .MoveLast
    End If
        .AddNew
    !usuario_liq = Usuario
    !id_pedido = Gcod_planilla
    !num_control = Gcod_control
    !num_factura = Gcod_factura
    
    !codigo = txt_codigo.Text
    !status = "VI"
    !fecha_pedido = Me.txt_fecha_pedido.Text
    !id_inst = Me.txt_cilindro.Text
    !monto_fac = CCur(Me.txt_total_cilindro.Text)
    !cant_pedido = CInt(Me.txt_cant.Text)
       .Update
                                                                                
   With clientes.Recordset
       !fecha_ult_pago = Me.txt_fecha_pedido.Text
          .Update
   End With
  End With
End If
 
 If txt_cilindro = "27" Then
        MsgBox "No Hay Instalaciones de 27 Kgs, por favor verifique ", vbInformation, "JerGas"
        Me.txt_cilindro.SetFocus
        Exit Sub
 End If
      
 With facturando.Recordset

    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
        .MoveLast
    End If
    .AddNew
    
    !usuario_liq = Usuario
    !id_pedido = Gcod_planilla
    !num_control = Gcod_control
    !num_factura = Gcod_factura
    
    !codigo = txt_codigo.Text
    !fecha_pedido = Me.txt_fecha_pedido.Text
    !cliente = Me.txt_cliente.Text
    !cedula = Me.txt_cedula.Text
    !direccion = Me.txt_direccion.Text
    !telefono_hab = Me.txt_telefono_hab.Text
    !observaciones = Me.txt_observaciones.Text
    !id_inst = Me.txt_cilindro.Text
    !descripcion = concepto
    !status = "VI"
    !id_ruta = Me.txt_ruta.Text
    !marca = "1"
    
    !cant_pedido = Me.txt_cant.Text
    !monto_fac = CCur(Me.txt_total_cilindro.Text)
    !iva = "0"
    !total_fac = CCur(Me.txt_total_cilindro.Text)

    .Update
End With
    
Gcod_planilla = FGNRO_LIQ()

With facturando.Recordset

       If Not (.BOF And .EOF) Then
         mvBookMark = .Bookmark
           .MoveLast
       End If
           .AddNew
    
       !usuario_liq = Usuario
       !id_pedido = Gcod_planilla
       !num_control = Gcod_control
       !num_factura = Gcod_factura
    
       !codigo = txt_codigo.Text
       !fecha_pedido = Me.txt_fecha_pedido.Text
       !cliente = Me.txt_cliente.Text
       !cedula = Me.txt_cedula.Text
       !direccion = Me.txt_direccion.Text
       !telefono_hab = Me.txt_telefono_hab.Text
       !observaciones = Me.txt_observaciones.Text
       !id_inst = Me.txt_cilindro.Text
       !descripcion = "Instalación para Suministro de GLP"
       !status = "VI"
       !id_ruta = Me.txt_ruta.Text
       !marca = "3"
       !monto_fac = CCur(Me.Text48.Text)
       !iva = CCur(Me.Text49.Text)
       !total_fac = CCur(Me.Text50.Text)
          .Update
   End With
   
   With nuevainstalaciones.Recordset
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
        .MoveLast
    End If
    .AddNew
    
    !id_pedido = Me.txt_pedidos.Text
    !num_control = Gcod_control
    !num_factura = Gcod_factura
    
    !codigo = txt_codigo.Text
    !fecha_pedido = Me.txt_fecha_pedido.Text
    !fecha_cancel = Me.txt_fecha_pedido.Text
    !id_inst = Me.txt_cilindro.Text
    !status = "CA"
    
    !cant_pedido = Me.txt_cant.Text
    !monto_fac = CCur(Me.txt_total_cilindro.Text)
    !monto_inst = CCur(Me.txt_precio_uni.Text)
    !sub_total = CCur(Me.Text48.Text)
    !tot_iva = CCur(Me.Text49.Text)
    !total_fac = CCur(Me.Text51.Text)
   
    .Update
End With
   
   
   Cancelar
         MsgBox "La Nueva Instalación Ha Sido Registrada, Presione Aceptar Para Finalizar ", vbInformation, "JerGas"
Unload Me
End Sub

Private Sub txt_cant_Change()

          If txt_cant.Text = "1" And txt_cilindro.Text = "10" Then
             With instalacion.Recordset
                mvBookMark = .Bookmark
                .MoveFirst
                .Find "id_inst = 10"
                
                precio = !precio_instalacion_1
                cant_cilindro = !precio_cilindro
                
                 cant_venta = CCur(precio - cant_cilindro)
                      paso1 = CCur(cant_venta / 1.09)
                      paso2 = Round(paso1, 2)
                      
                 total_iva = CCur(paso2 * !alicuota) / 100
                 paso3 = Round(total_iva, 2)
                
                 Subtotal = CCur(paso2 + cant_cilindro)
                
                Me.txt_precio_uni.Text = CCur(paso2)
                Me.txt_total_cilindro.Text = CCur(cant_cilindro)
                Me.txt_monto_fac.Text = CCur(Subtotal)
                Me.txt_iva.Text = paso3
                Me.txt_total_pagar.Text = CCur(Subtotal + paso3)
                pasoFinal = CCur(Subtotal + paso3)
                Me.Text51.Text = pasoFinal
             End With
          End If
            
  
          If txt_cant.Text = "2" And txt_cilindro.Text = "10" Then
             With instalacion.Recordset
                mvBookMark = .Bookmark
                .MoveFirst
                .Find "id_inst = 10"
                
                precio = !precio_instalacion_2
                cant_cilindro = !precio_cilindro * 2
                
                 cant_venta = CCur(precio - cant_cilindro)
                      paso1 = CCur(cant_venta / 1.09)
                      paso2 = Round(paso1, 2)
                      
                 total_iva = CCur(paso2 * !alicuota) / 100
                 paso3 = Round(total_iva, 2)
                
                 Subtotal = CCur(paso2 + cant_cilindro)
                
                Me.txt_precio_uni.Text = CCur(paso2)
                Me.txt_total_cilindro.Text = CCur(cant_cilindro)
                Me.txt_monto_fac.Text = CCur(Subtotal)
                Me.txt_iva.Text = paso3
                Me.txt_total_pagar.Text = CCur(Subtotal + paso3)
                pasoFinal = CCur(Subtotal + paso3)
                Me.Text51.Text = pasoFinal
             End With
          End If
          
          If txt_cant.Text = "1" And txt_cilindro.Text = "18" Then
             With instalacion.Recordset
                mvBookMark = .Bookmark
                .MoveFirst
                .Find "id_inst = 18"
                
                precio = !precio_instalacion_1
                cant_cilindro = !precio_cilindro
                
                 cant_venta = CCur(precio - cant_cilindro)
                      paso1 = CCur(cant_venta / 1.09)
                      paso2 = Round(paso1, 2)
                      
                 total_iva = CCur(paso2 * !alicuota) / 100
                 paso3 = Round(total_iva, 2)
                
                 Subtotal = CCur(paso2 + cant_cilindro)
                
                Me.txt_precio_uni.Text = CCur(paso2)
                Me.txt_total_cilindro.Text = CCur(cant_cilindro)
                Me.txt_monto_fac.Text = CCur(Subtotal)
                Me.txt_iva.Text = paso3
                Me.txt_total_pagar.Text = CCur(Subtotal + paso3)
                pasoFinal = CCur(Subtotal + paso3)
                Me.Text51.Text = pasoFinal
             End With
          End If
            
          If txt_cant.Text = "2" And txt_cilindro.Text = "18" Then
             With instalacion.Recordset
                mvBookMark = .Bookmark
                .MoveFirst
                .Find "id_inst = 18"
                
                precio = !precio_instalacion_2
                cant_cilindro = !precio_cilindro * 2
                
                 cant_venta = CCur(precio - cant_cilindro)
                      paso1 = CCur(cant_venta / 1.09)
                      paso2 = Round(paso1, 2)
                      
                 total_iva = CCur(paso2 * !alicuota) / 100
                 paso3 = Round(total_iva, 2)
                
                 Subtotal = CCur(paso2 + cant_cilindro)
                
                Me.txt_precio_uni.Text = CCur(paso2)
                Me.txt_total_cilindro.Text = CCur(cant_cilindro)
                Me.txt_monto_fac.Text = CCur(Subtotal)
                Me.txt_iva.Text = paso3
                Me.txt_total_pagar.Text = CCur(Subtotal + paso3)
                pasoFinal = CCur(Subtotal + paso3)
                Me.Text51.Text = pasoFinal
             End With
          End If
          
          If txt_cant.Text = "1" And txt_cilindro.Text = "43" Then
             With instalacion.Recordset
                mvBookMark = .Bookmark
                .MoveFirst
                .Find "id_inst = 43"
                
                precio = !precio_instalacion_1
                cant_cilindro = !precio_cilindro
                
                 cant_venta = CCur(precio - cant_cilindro)
                      paso1 = CCur(cant_venta / 1.09)
                      paso2 = Round(paso1, 2)
                      
                 total_iva = CCur(paso2 * !alicuota) / 100
                 paso3 = Round(total_iva, 2)
                
                 Subtotal = CCur(paso2 + cant_cilindro)
                
                Me.txt_precio_uni.Text = CCur(paso2)
                Me.txt_total_cilindro.Text = CCur(cant_cilindro)
                Me.txt_monto_fac.Text = CCur(Subtotal)
                Me.txt_iva.Text = paso3
                Me.txt_total_pagar.Text = CCur(Subtotal + paso3)
                pasoFinal = CCur(Subtotal + paso3)
                Me.Text51.Text = pasoFinal
             End With
          End If
            
          If txt_cant.Text = "2" And txt_cilindro.Text = "43" Then
             With instalacion.Recordset
                mvBookMark = .Bookmark
                .MoveFirst
                .Find "id_inst = 43"
                
                precio = !precio_instalacion_2
                cant_cilindro = !precio_cilindro * 2
                
                 cant_venta = CCur(precio - cant_cilindro)
                      paso1 = CCur(cant_venta / 1.09)
                      paso2 = Round(paso1, 2)
                      
                 total_iva = CCur(paso2 * !alicuota) / 100
                 paso3 = Round(total_iva, 2)
                
                 Subtotal = CCur(paso2 + cant_cilindro)
                
                Me.txt_precio_uni.Text = CCur(paso2)
                Me.txt_total_cilindro.Text = CCur(cant_cilindro)
                Me.txt_monto_fac.Text = CCur(Subtotal)
                Me.txt_iva.Text = paso3
                Me.txt_total_pagar.Text = CCur(Subtotal + paso3)
                pasoFinal = CCur(Subtotal + paso3)
                Me.Text51.Text = pasoFinal
             End With
          End If
 
   acumula1 = paso2
   acumula2 = paso3
   acumula3 = cant_venta
   
   Me.Text48.Text = acumula1
   Me.Text49.Text = acumula2
   Me.Text50.Text = acumula3

 End Sub

Private Sub txt_cant_a_LostFocus()

Check4.Value = 0
'
'   acumula1 = acumula1 + paso2A
'   acumula2 = acumula2 + paso3A
'   acumula3 = acumula3 + pasoprevioA
'
'   Me.Text48.Text = acumula1
'   Me.Text49.Text = acumula2
'   Me.Text50.Text = acumula3
'   Me.Text51.Text = txt_monto_general.Text
'
End Sub
Private Sub txt_cant_a_Change()

      If txt_cant_a.Text <> "" Then
             
         With materiales.Recordset
            materiales.Recordset.MoveFirst
            strquery = "descripcion = '" & Me.txt_descripcion.Text & "'"
            materiales.Recordset.Find strquery

            precio = !precio_venta
                 impuesto = !iva
                
                 cant_ventaA = Me.txt_cant_a.Text * precio
                      paso1A = CCur(cant_ventaA / 1.09)
                      paso2A = Round(paso1A, 2)
                      
                 total_ivaA = CCur(paso2A * impuesto) / 100
                 paso3A = Round(total_ivaA, 2)
                   
                 precio_unitarioA = paso2A

                 Me.txt_precio_a.Text = CCur(paso2A)
                 Me.txt_iva_a.Text = CCur(paso3A)
                 Me.txt_total_a.Text = CCur(paso2A + paso3A)
                 pasoprevioA = cant_ventaA
                 pasoFinalA = CCur(paso2A + paso3A)
'         End With
'      End If
      
    Frame_final.Visible = True
    
    uno = paso3
    dos = paso3A
          iva_final = uno + dos
    
    tres = cant_venta
    cuatro = cant_ventaA
          monto_general = tres + cuatro
          
    cinco = paso2
    seis = paso2A
          total_precio_inst = cinco + seis
    
    Me.txt_iva_general.Text = iva_final
    Me.Text49.Text = iva_final
    
    Me.txt_monto_general.Text = (pasoFinal + pasoFinalA)
    
    Me.Text50.Text = monto_general
    Me.Text51.Text = (pasoFinal + pasoFinalA)

Check4.Visible = True

   acumula1 = acumula1 + paso2A
   acumula2 = acumula2 + paso3A
   acumula3 = acumula3 + pasoprevioA
   
   Me.Text48.Text = acumula1
   Me.Text49.Text = acumula2
   Me.Text50.Text = acumula3
   Me.Text51.Text = txt_monto_general.Text
End With
      End If
End Sub

Private Sub txt_cliente_LostFocus()

Dim monto1 As Integer

Dim fec As Date
Dim ano As Date
Dim strquery As String
Dim resta As Integer
Dim Total As Integer
Dim bandera As Boolean
Dim abc, ncliente, contrato

On Error GoTo ControlError
    
    If IsNull(txt_cliente.Text) Or txt_cliente.Text = "" Then
        MsgBox "Apellido y Nombre no puede ser nulo, por favor verifique ", vbInformation, "JerGas"
        Me.txt_cliente.SetFocus
        Exit Sub
    End If
    
    If IsNull(Me.txt_ruta.Text) Or txt_ruta.Text = "" Then
        MsgBox "La ruta no puede ser nulo, por favor verifique ", vbInformation, "JerGas"
        Me.txt_ruta.SetFocus
        Exit Sub
    End If
    
    If IsNull(Me.txt_cilindro.Text) Or txt_cilindro.Text = "" Then
        MsgBox "Debe seleccionar un tipo de cilindro, por favor verifique ", vbInformation, "JerGas"
        Me.txt_cilindro.SetFocus
        Exit Sub
    End If
    
    'Buscamos el ultimo numero de cliente
    
    control_clientes.Recordset.MoveFirst
    abc = Left(txt_cliente.Text, 1)
    strquery = "id = '" & abc & "'"
    control_clientes.Recordset.Find strquery
    If control_clientes.Recordset.EOF Then
             MsgBox "En el campo Apellido y Nombre la primera caracter suministrado debe ser una letra, por favor verifique ", vbInformation, "JerGas"
             Me.txt_cliente.SetFocus
            Exit Sub
    Else
        ncliente = CInt(control_clientes.Recordset!valor) + 1
        ncliente = Format(ncliente, "0000")
        txt_codigo.Text = "" + abc + "-" + Me.txt_ruta.Text + "-" + ncliente + ""
        control_clientes.Recordset!valor = ncliente
        control_clientes.Recordset.Update
    End If
      
    'Buscamos la instalacion con respecto a la bonbona seleccionada
    'Buscamos el ultimo numero de cliente
    
    instalacion.Recordset.MoveFirst
    strquery = "id_inst = '" & Me.txt_cilindro.Text & "'"
    instalacion.Recordset.Find strquery
     If instalacion.Recordset.EOF Then
              MsgBox "Verifique la bombona suministrada ", vbInformation, "Jergas"
              Me.txt_cilindro.SetFocus
            Exit Sub
    Else
    
        contrato = CInt(instalacion.Recordset!ncontrato) + 1
        contrato = Format(contrato, "00000")
        Me.txt_contrato.Text = "" + Me.txt_cilindro.Text + "-" + contrato + ""
        instalacion.Recordset!ncontrato = contrato
       instalacion.Recordset.Update
      End If
    
    With clientes.Recordset
        mvBookMark = .Bookmark
        .Update
        .Bookmark = mvBookMark
    End With

  
  If txt_cilindro = "10" Then

     Me.kit.CommandType = adCmdText
     Me.kit.RecordSource = "select * from tbl_kit WHERE tbl_kit.id_inst = 10 ORDER BY id_inst ASC"
     Me.kit.Refresh


     Label11.Caption = "Cilindros de 10 Kgs"
     Text5.Text = 1
      Label1.Caption = "Cilindros de 10 Kgs"
     Text6.Text = 1
      Label2.Caption = "Regulador de 10 Kgs"
     Text46.Text = 0
      Label14.Caption = "Rabo de Cochino"
     Text7.Text = 0
      Label15.Caption = "Conector de Regulador "
     Text8.Text = 0
      Label4.Caption = "Copas de 3/8 Pul"
     Text9.Text = 0
      Label5.Caption = "Reductor 1/2 x 3/8 Pul"
     Text10.Text = 0
      Label6.Caption = "Teek Cheek"
     Text11.Text = 0
      Label7.Caption = "Mts de Tubo de Cobre de 3/8 Pul"

  Activar

                    Text14.Text = Format(Text14.Text, "standard")
                    Text15.Text = Format(Text15.Text, "standard")
                    Text16.Text = Format(Text16.Text, "standard")
                    Text17.Text = Format(Text17.Text, "standard")
                    Text18.Text = Format(Text18.Text, "standard")
                    Text19.Text = Format(Text19.Text, "standard")
                    Text20.Text = Format(Text20.Text, "standard")
                    Text45.Text = Format(Text45.Text, "standard")

               Text22.Text = CSng(Text14.Text) * CSng(Text5.Text)
               txt_iva1.Text = (CSng(Text22.Text) * 9) / 100
                    Text22.Text = Format(Text22.Text, "standard")
               Text23.Text = CSng(Text15.Text) * CSng(Text6.Text)
               txt_iva2.Text = (CSng(Text23.Text) * 9) / 100
                    Text23.Text = Format(Text23.Text, "standard")
               Text24.Text = CSng(Text16.Text) * CSng(Text7.Text)
               txt_iva3.Text = (CSng(Text24.Text) * 9) / 100
                    Text24.Text = Format(Text24.Text, "standard")
               Text25.Text = CSng(Text17.Text) * CSng(Text8.Text)
               txt_iva4.Text = (CSng(Text25.Text) * 9) / 100
                    Text25.Text = Format(Text25.Text, "standard")
               Text26.Text = CSng(Text18.Text) * CSng(Text9.Text)
               txt_iva5.Text = (CSng(Text26.Text) * 9) / 100
                    Text26.Text = Format(Text26.Text, "standard")
               Text27.Text = CSng(Text19.Text) * CSng(Text10.Text)
               txt_iva6.Text = (CSng(Text27.Text) * 9) / 100
                    Text27.Text = Format(Text27.Text, "standard")
               Text28.Text = CSng(Text20.Text) * CSng(Text11.Text)
               txt_iva7.Text = (CSng(Text28.Text) * 9) / 100
                    Text28.Text = Format(Text28.Text, "standard")
               Text44.Text = CSng(Text46.Text) * CSng(Text45.Text)
               txt_iva8.Text = (CSng(Text46.Text) * 9) / 100
                    Text44.Text = Format(Text44.Text, "standard")


      Text32.Text = CSng(txt_iva1.Text) + CSng(txt_iva2.Text) + CSng(txt_iva3.Text) + CSng(txt_iva4.Text) + CSng(txt_iva5.Text) + CSng(txt_iva6.Text) + CSng(txt_iva7.Text) + CSng(txt_iva8.Text)
      Text21.Text = CSng(Text22.Text) + CSng(Text23.Text) + CSng(Text24.Text) + CSng(Text25.Text) + CSng(Text26.Text) + CSng(Text27.Text) + CSng(Text28.Text) + CSng(Text44.Text) + CSng(txt_iva1.Text) + CSng(txt_iva2.Text) + CSng(txt_iva3.Text) + CSng(txt_iva4.Text) + CSng(txt_iva5.Text) + CSng(txt_iva6.Text) + CSng(txt_iva7.Text) + CSng(txt_iva8.Text)
                    Text21.Text = Format(Text21.Text, "standard")

End If

If txt_cilindro = "18" Then
     
     Me.kit.CommandType = adCmdText
     Me.kit.RecordSource = "select * from tbl_kit WHERE tbl_kit.id_inst = 18 ORDER BY id_inst ASC"
     Me.kit.Refresh

     Label11.Caption = "Cilindros de 18 Kgs"
     Text5.Text = 2
      Label1.Caption = "Cilindros de 18 Kgs"
     Text6.Text = 1
      Label2.Caption = "Regulador de 18 Kgs"
     Text46.Text = 2
      Label14.Caption = "Rabo de Cochino"
     Text7.Text = 1
      Label15.Caption = "Conector de Regulador "
     Text8.Text = 1
      Label4.Caption = "Copas de 3/8 Pul"
     Text9.Text = 2
      Label5.Caption = "Reductor 1/2 x 3/8 Pul"
     Text10.Text = 1
      Label6.Caption = "Teek Cheek"
     Text11.Text = 2
      Label7.Caption = "Mts de Tubo de Cobre de 3/8 Pul"

  Activar

                    Text14.Text = Format(Text14.Text, "standard")
                    Text15.Text = Format(Text15.Text, "standard")
                    Text16.Text = Format(Text16.Text, "standard")
                    Text17.Text = Format(Text17.Text, "standard")
                    Text18.Text = Format(Text18.Text, "standard")
                    Text19.Text = Format(Text19.Text, "standard")
                    Text20.Text = Format(Text20.Text, "standard")
                    Text45.Text = Format(Text45.Text, "standard")

               Text22.Text = CSng(Text14.Text) * CSng(Text5.Text)
               txt_iva1.Text = (CSng(Text22.Text) * 9) / 100
                    Text22.Text = Format(Text22.Text, "standard")
               Text23.Text = CSng(Text15.Text) * CSng(Text6.Text)
               txt_iva2.Text = (CSng(Text23.Text) * 9) / 100
                    Text23.Text = Format(Text23.Text, "standard")
               Text24.Text = CSng(Text16.Text) * CSng(Text7.Text)
               txt_iva3.Text = (CSng(Text24.Text) * 9) / 100
                    Text24.Text = Format(Text24.Text, "standard")
               Text25.Text = CSng(Text17.Text) * CSng(Text8.Text)
               txt_iva4.Text = (CSng(Text25.Text) * 9) / 100
                    Text25.Text = Format(Text25.Text, "standard")
               Text26.Text = CSng(Text18.Text) * CSng(Text9.Text)
               txt_iva5.Text = (CSng(Text26.Text) * 9) / 100
                    Text26.Text = Format(Text26.Text, "standard")
               Text27.Text = CSng(Text19.Text) * CSng(Text10.Text)
               txt_iva6.Text = (CSng(Text27.Text) * 9) / 100
                    Text27.Text = Format(Text27.Text, "standard")
               Text28.Text = CSng(Text20.Text) * CSng(Text11.Text)
               txt_iva7.Text = (CSng(Text28.Text) * 9) / 100
                    Text28.Text = Format(Text28.Text, "standard")
               Text44.Text = CSng(Text46.Text) * CSng(Text45.Text)
               txt_iva8.Text = (CSng(Text46.Text) * 9) / 100
                    Text44.Text = Format(Text44.Text, "standard")


      Text32.Text = CSng(txt_iva1.Text) + CSng(txt_iva2.Text) + CSng(txt_iva3.Text) + CSng(txt_iva4.Text) + CSng(txt_iva5.Text) + CSng(txt_iva6.Text) + CSng(txt_iva7.Text) + CSng(txt_iva8.Text)
      Text21.Text = CSng(Text22.Text) + CSng(Text23.Text) + CSng(Text24.Text) + CSng(Text25.Text) + CSng(Text26.Text) + CSng(Text27.Text) + CSng(Text28.Text) + CSng(Text44.Text) + CSng(txt_iva1.Text) + CSng(txt_iva2.Text) + CSng(txt_iva3.Text) + CSng(txt_iva4.Text) + CSng(txt_iva5.Text) + CSng(txt_iva6.Text) + CSng(txt_iva7.Text) + CSng(txt_iva8.Text)
                    Text21.Text = Format(Text21.Text, "standard")

End If

If txt_cilindro = "27" Then
        
        MsgBox "No Hay Instalaciones de 27 Kgs, por favor verifique ", vbInformation, "JerGas"
        Me.txt_ruta.SetFocus
        Exit Sub
End If

If txt_cilindro = "43" Then
     
     Me.kit.CommandType = adCmdText
     Me.kit.RecordSource = "select * from tbl_kit WHERE tbl_kit.id_inst = 43 ORDER BY id_inst ASC"
     Me.kit.Refresh

      Label11.Caption = "Cilindros de 43 Kgs"
     Text5.Text = 2
      Label1.Caption = "Cilindros de 43 Kgs"
     Text6.Text = 1
      Label2.Caption = "Regulador de 43 Kgs"
     Text46.Text = 2
      Label14.Caption = "Rabo de Cochino"
     Text7.Text = 1
      Label15.Caption = "Conector de Regulador "
     Text8.Text = 1
      Label4.Caption = "Copas de 3/8 Pul"
     Text9.Text = 2
      Label5.Caption = "Reductor 1/2 x 3/8 Pul"
     Text10.Text = 1
      Label6.Caption = "Teek Cheek"
     Text11.Text = 2
      Label7.Caption = "Mts de Tubo de Cobre de 3/8 Pul"

  Activar

                    Text14.Text = Format(Text14.Text, "standard")
                    Text15.Text = Format(Text15.Text, "standard")
                    Text16.Text = Format(Text16.Text, "standard")
                    Text17.Text = Format(Text17.Text, "standard")
                    Text18.Text = Format(Text18.Text, "standard")
                    Text19.Text = Format(Text19.Text, "standard")
                    Text20.Text = Format(Text20.Text, "standard")
                    Text45.Text = Format(Text45.Text, "standard")

               Text22.Text = CSng(Text14.Text) * CSng(Text5.Text)
               txt_iva1.Text = (CSng(Text22.Text) * 9) / 100
                    Text22.Text = Format(Text22.Text, "standard")
               Text23.Text = CSng(Text15.Text) * CSng(Text6.Text)
               txt_iva2.Text = (CSng(Text23.Text) * 9) / 100
                    Text23.Text = Format(Text23.Text, "standard")
               Text24.Text = CSng(Text16.Text) * CSng(Text7.Text)
               txt_iva3.Text = (CSng(Text24.Text) * 9) / 100
                    Text24.Text = Format(Text24.Text, "standard")
               Text25.Text = CSng(Text17.Text) * CSng(Text8.Text)
               txt_iva4.Text = (CSng(Text25.Text) * 9) / 100
                    Text25.Text = Format(Text25.Text, "standard")
               Text26.Text = CSng(Text18.Text) * CSng(Text9.Text)
               txt_iva5.Text = (CSng(Text26.Text) * 9) / 100
                    Text26.Text = Format(Text26.Text, "standard")
               Text27.Text = CSng(Text19.Text) * CSng(Text10.Text)
               txt_iva6.Text = (CSng(Text27.Text) * 9) / 100
                    Text27.Text = Format(Text27.Text, "standard")
               Text28.Text = CSng(Text20.Text) * CSng(Text11.Text)
               txt_iva7.Text = (CSng(Text28.Text) * 9) / 100
                    Text28.Text = Format(Text28.Text, "standard")
               Text44.Text = CSng(Text46.Text) * CSng(Text45.Text)
               txt_iva8.Text = (CSng(Text46.Text) * 9) / 100
                    Text44.Text = Format(Text44.Text, "standard")


      Text32.Text = CSng(txt_iva1.Text) + CSng(txt_iva2.Text) + CSng(txt_iva3.Text) + CSng(txt_iva4.Text) + CSng(txt_iva5.Text) + CSng(txt_iva6.Text) + CSng(txt_iva7.Text) + CSng(txt_iva8.Text)
      Text21.Text = CSng(Text22.Text) + CSng(Text23.Text) + CSng(Text24.Text) + CSng(Text25.Text) + CSng(Text26.Text) + CSng(Text27.Text) + CSng(Text28.Text) + CSng(Text44.Text) + CSng(txt_iva1.Text) + CSng(txt_iva2.Text) + CSng(txt_iva3.Text) + CSng(txt_iva4.Text) + CSng(txt_iva5.Text) + CSng(txt_iva6.Text) + CSng(txt_iva7.Text) + CSng(txt_iva8.Text)
                    Text21.Text = Format(Text21.Text, "standard")
End If

     Exit Sub    ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "Control del Cliente")
        Case 3314
            MsgBox "Verifique la Cédula ", vbOKOnly, "Control del Cliente"
        Case 524
            MsgBox "Verifique el Nombre y Apellido  ", vbOKOnly, "Control del Cliente"
        Case -2147467259
            MsgBox "Error, la cédula suministrada ya existe", vbOKOnly, "Control del Cliente"
        Case -2147217842
            MsgBox "Error, cancele la operación y vuelva a intentarlo", vbOKOnly, "Control del Cliente"
        Case -2147217887
            MsgBox "Error, al crear histórico, se recomienda borrar el registro y volverlo a crear", vbOKOnly, "Control del Cliente"
    End Select

End Sub

Private Sub Command1_Click()

  On Error GoTo ControlError

                    Text14.Text = Format(Text14.Text, "standard")
                    Text15.Text = Format(Text15.Text, "standard")
                    Text16.Text = Format(Text16.Text, "standard")
                    Text17.Text = Format(Text17.Text, "standard")
                    Text18.Text = Format(Text18.Text, "standard")
                    Text19.Text = Format(Text19.Text, "standard")
                    Text20.Text = Format(Text20.Text, "standard")
               Text22.Text = CSng(Text14.Text) * CSng(Text5.Text)
               txt_iva1.Text = (CSng(Text22.Text) * 9) / 100
                    Text22.Text = Format(Text22.Text, "standard")
               Text23.Text = CSng(Text15.Text) * CSng(Text6.Text)
               txt_iva2.Text = (CSng(Text23.Text) * 9) / 100
                    Text23.Text = Format(Text23.Text, "standard")
               Text24.Text = CSng(Text16.Text) * CSng(Text7.Text)
               txt_iva3.Text = (CSng(Text24.Text) * 9) / 100
                    Text24.Text = Format(Text24.Text, "standard")
               Text25.Text = CSng(Text17.Text) * CSng(Text8.Text)
               txt_iva4.Text = (CSng(Text25.Text) * 9) / 100
                    Text25.Text = Format(Text25.Text, "standard")
               Text26.Text = CSng(Text18.Text) * CSng(Text9.Text)
               txt_iva5.Text = (CSng(Text26.Text) * 9) / 100
                    Text26.Text = Format(Text26.Text, "standard")
               Text27.Text = CSng(Text19.Text) * CSng(Text10.Text)
               txt_iva6.Text = (CSng(Text27.Text) * 9) / 100
                    Text27.Text = Format(Text27.Text, "standard")
               Text28.Text = CSng(Text20.Text) * CSng(Text11.Text)
               txt_iva7.Text = (CSng(Text28.Text) * 9) / 100
                    Text28.Text = Format(Text28.Text, "standard")
               txt_iva8.Text = (CSng(Text45.Text) * 9) / 100
                    Text44.Text = Format(Text44.Text, "standard")

      Text32.Text = CSng(txt_iva1.Text) + CSng(txt_iva2.Text) + CSng(txt_iva3.Text) + CSng(txt_iva4.Text) + CSng(txt_iva5.Text) + CSng(txt_iva6.Text) + CSng(txt_iva7.Text) + CSng(txt_iva8.Text)
      Text21.Text = CSng(Text22.Text) + CSng(Text23.Text) + CSng(Text24.Text) + CSng(Text25.Text) + CSng(Text26.Text) + CSng(Text27.Text) + CSng(Text28.Text) + CSng(Text44.Text) + CSng(txt_iva1.Text) + CSng(txt_iva2.Text) + CSng(txt_iva3.Text) + CSng(txt_iva4.Text) + CSng(txt_iva5.Text) + CSng(txt_iva6.Text) + CSng(txt_iva7.Text) + CSng(txt_iva8.Text)
                    Text21.Text = Format(Text21.Text, "standard")

Check1.Value = 0
Command1.Enabled = False

ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("DEBE INGRESAR (0) SI NO HAY NINGUN VALOR, POR FAVOR VERIFIQUE", vbOKOnly, "Control de Clientes")
    End Select
End Sub

Private Sub Check1_Click()

   If Check1 = 1 Then
     Command1.Enabled = True
     Text5.Locked = False
     Text6.Locked = False
     Text7.Locked = False
     Text8.Locked = False
     Text9.Locked = False
     Text10.Locked = False
     Text11.Locked = False
     Text46.Locked = False
   End If
End Sub

Private Sub cmdagregar_Click()
On Error GoTo AddErr
    DTPicker1.Value = Date
    
    Dim fec As Date
    Dim ano As Date
    
    cmdagregar.Visible = False
    cmdcancelar.Enabled = True
    cmdeliminar.Enabled = True
    cmdsalir.Enabled = False
    
    cmdguardar.Enabled = True
    cmdcancelar.Visible = True
  
    DTPicker1.Enabled = True
    Text12.Enabled = True
    txt_cliente.Locked = False
    txt_cedula.Locked = False
    txt_direccion.Locked = False
    txt_telefono_hab.Locked = False
    txt_telefono_cel.Locked = False
    txt_observaciones.Locked = False
    txt_correo.Locked = False
    txt_ruta.Locked = False
    txt_status.Locked = False
    txt_cilindro.Locked = False
    
    DTPicker1.Enabled = True
    Text12.Enabled = True
    txt_cliente.Enabled = True
    txt_cedula.Enabled = True
    txt_direccion.Enabled = True
    txt_telefono_hab.Enabled = True
    txt_telefono_cel.Enabled = True
    txt_observaciones.Enabled = True
    txt_correo.Enabled = True
    txt_ruta.Enabled = True
    txt_status.Enabled = True
    txt_cilindro.Enabled = True
    
    fecha = DTPicker1.Value
    Text1.Text = DateAdd("m", 1, fec)
    Text2.Text = fecha
    Text12.Text = fecha
    Text2.Text = DateAdd("yyyy", 1, ano)
    
  With clientes.Recordset
       If Not (.BOF And .EOF) Then
         mvBookMark = .Bookmark
           .MoveLast
       End If
          .AddNew
  End With
    
    Me.DTPicker1.Value = Date
    Me.txt_status.Text = "VI"
    Me.Text12.Text = Date
  
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdcancelar_Click()

On Error GoTo ControlError
    
    clientes.Recordset.CancelUpdate
            cmdguardar.Enabled = False
            cmdagregar.Visible = True
            cmdsalir.Enabled = True
            cmdeliminar.Enabled = False
            cmdcancelar.Visible = False
            txt_cliente.Locked = True
            txt_cedula.Locked = True
            txt_direccion.Locked = True
            txt_telefono_hab.Locked = True
            txt_telefono_cel.Locked = True
            txt_observaciones.Locked = True
            txt_correo.Locked = True
            txt_ruta.Locked = True
            txt_status.Locked = True
            txt_cilindro.Locked = True
    
    txt_codigo.Text = ""
    txt_contrato.Text = ""
    txt_cedula.Text = ""
    txt_cliente.Text = ""
    txt_direccion.Text = ""
    txt_telefono_hab.Text = ""
    txt_telefono_cel.Text = ""
    txt_correo.Text = ""
    txt_ruta.Text = ""
    txt_cilindro.Text = ""
    txt_status.Text = ""
    txt_observaciones.Text = ""
    txt_cant.Text = ""
    txt_instalacion.Text = ""
    txt_precio_uni.Text = ""
    txt_iva.Text = ""
    txt_monto_fac.Text = ""
    txt_total_cilindro.Text = ""
    txt_total_pagar.Text = ""
    
    Frame_total.Visible = False
    cmd_imprime.Visible = False
    Frame1.Visible = False
    Frame_final.Visible = False
    Check3.Value = 0
    Frame_adicional.Visible = False
    
    Label1.Visible = False
    Label2.Visible = False
    Label3.Visible = False
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Label7.Visible = False
    Label8.Visible = False
    Label9.Visible = False
    Label10.Visible = False
    Label11.Visible = False
    
    Text5.Visible = False
    Text6.Visible = False
    Text7.Visible = False
    Text8.Visible = False
    Text9.Visible = False
    Text10.Visible = False
    Text11.Visible = False
    Text14.Visible = False
    Text15.Visible = False
    Text16.Visible = False
    Text17.Visible = False
    Text18.Visible = False
    Text19.Visible = False
    Text20.Visible = False
    Text21.Visible = False
    Text22.Visible = False
    Text23.Visible = False
    Text24.Visible = False
    Text25.Visible = False
    Text26.Visible = False
    Text27.Visible = False
    Text28.Visible = False
    
    Label12.Visible = False
    txt_iva1.Visible = False
    txt_iva2.Visible = False
    txt_iva3.Visible = False
    txt_iva4.Visible = False
    txt_iva5.Visible = False
    txt_iva6.Visible = False
    txt_iva7.Visible = False
    Cancelar
    Exit Sub    ' Salir para evitar el controlador.

ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "Control de Clientes")
    End Select
End Sub

Private Sub Form_Load()
'Call actualizar_cn("SQL Server")
    cmdcancelar.Visible = False
    cmdeliminar.Enabled = False
      Me.Text5.Text = Me.Text33.Text
    DTPicker1.Enabled = True
    clientes.Refresh
    kit.Refresh
    Text1.Text = Date
    Me.txt_fecha_pedido = Date
    DTPicker1.Enabled = False
    Text12.Enabled = False
    txt_cliente.Enabled = False
    txt_cedula.Enabled = False
    txt_direccion.Enabled = False
    txt_telefono_hab.Enabled = False
    txt_telefono_cel.Enabled = False
    txt_observaciones.Enabled = False
    txt_correo.Enabled = False
    txt_ruta.Enabled = False
    txt_status.Enabled = False
    txt_cilindro.Enabled = False
    txt_instalacion.Text = ""
End Sub

Private Sub cmdsalir_Click()
  Unload Me
End Sub

Private Sub cmdeliminar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdeliminar.FontBold = True
Me.cmdguardar.FontBold = False
Me.cmdagregar.FontBold = False
Me.cmdcancelar.FontBold = False
Me.cmdsalir.FontBold = False
End Sub

Private Sub cmdguardar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdguardar.FontBold = True
Me.cmdagregar.FontBold = False
Me.cmdcancelar.FontBold = False
Me.cmdsalir.FontBold = False
Me.cmdeliminar.FontBold = False
End Sub

Private Sub cmdcancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdcancelar.FontBold = True
Me.cmdagregar.FontBold = False
Me.cmdguardar.FontBold = False
Me.cmdsalir.FontBold = False
Me.cmdeliminar.FontBold = False
End Sub

Private Sub cmdagregar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdagregar.FontBold = True
Me.cmdguardar.FontBold = False
Me.cmdcancelar.FontBold = False
Me.cmdsalir.FontBold = False
Me.cmdeliminar.FontBold = False
End Sub

Private Sub cmdsalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdsalir.FontBold = True
Me.cmdguardar.FontBold = False
Me.cmdcancelar.FontBold = False
Me.cmdagregar.FontBold = False
Me.cmdeliminar.FontBold = False
End Sub

Private Sub txt_cilindro_Click(Area As Integer)
  
  If txt_cilindro.Text = "10" Then
    Me.txt_instalacion.Text = "Cilindros de 10 Kgs"
  End If
  
  If txt_cilindro.Text = "18" Then
    Me.txt_instalacion.Text = "Cilindros de 18 Kgs"
  End If
  
  If txt_cilindro.Text = "27" Then
    Me.txt_instalacion.Text = "Cilindros de 27 Kgs"
  End If
  
  If txt_cilindro.Text = "43" Then
    Me.txt_instalacion.Text = "Cilindros de 43 Kgs"
  End If
  
End Sub

Private Sub txt_telefono_hab_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_telefono_cel_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_correo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Index <> 6 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 13 Or Index = 6 Or Index = 9 Then
        If ((KeyAscii < 48) Or (KeyAscii > 57)) And (KeyAscii <> 8) Then KeyAscii = 0
    End If
End Sub

Private Sub txt_cedula_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Index <> 6 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 13 Or Index = 6 Or Index = 9 Then
        If ((KeyAscii < 48) Or (KeyAscii > 57)) And (KeyAscii <> 8) Then KeyAscii = 0
    End If
End Sub

Private Sub txt_direccion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Index <> 6 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 13 Or Index = 6 Or Index = 9 Then
        If ((KeyAscii < 48) Or (KeyAscii > 57)) And (KeyAscii <> 8) Then KeyAscii = 0
    End If
End Sub

Private Sub txt_observaciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Index <> 6 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 13 Or Index = 6 Or Index = 9 Then
        If ((KeyAscii < 48) Or (KeyAscii > 57)) And (KeyAscii <> 8) Then KeyAscii = 0
   End If
End Sub

Private Sub txt_ruta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub text6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub text7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub text8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub text9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub text10_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub text11_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_cant_a_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub


Sub Cancelar()
    'Procedimiento para blanquear el formulario
    txt_codigo.Text = ""
    txt_contrato.Text = ""
    txt_cedula.Text = ""
    txt_cliente.Text = ""
    txt_direccion.Text = ""
    txt_telefono_hab.Text = ""
    txt_telefono_cel.Text = ""
    txt_correo.Text = ""
    txt_ruta.Text = ""
    txt_cilindro.Text = ""
    txt_status.Text = ""
    txt_observaciones.Text = ""
    txt_cant.Text = ""
    txt_instalacion.Text = ""
    txt_precio_uni.Text = ""
    txt_iva.Text = ""
    txt_monto_fac.Text = ""
    txt_total_cilindro.Text = ""
    txt_total_pagar.Text = ""
    
    cmdagregar.Visible = True
    cmdsalir.Enabled = True
    cmdcancelar.Visible = False
    cmdguardar.Enabled = False
    cmdagregar.SetFocus
    cmdeliminar.Enabled = True

    Frame_kit.Enabled = False
    Frame_total.Visible = False
    cmd_imprime.Visible = False
    Frame1.Visible = False
    Command1.Visible = False
    Check1.Visible = False
    Frame_adicional.Visible = False
    Check3.Value = 0
    
    Label1.Visible = False
    Label2.Visible = False
    Label3.Visible = False
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Label7.Visible = False
    Label8.Visible = False
    Label9.Visible = False
    Label10.Visible = False
    Label11.Visible = False
    Label12.Visible = False
    Label14.Visible = False
    Label15.Visible = False
    
    txt_iva1.Visible = False
    txt_iva2.Visible = False
    txt_iva3.Visible = False
    txt_iva4.Visible = False
    txt_iva5.Visible = False
    txt_iva6.Visible = False
    txt_iva7.Visible = False
    txt_iva8.Visible = False

    Text5.Visible = False
    Text6.Visible = False
    Text7.Visible = False
    Text8.Visible = False
    Text9.Visible = False
    Text10.Visible = False
    Text11.Visible = False
    Text14.Visible = False
    Text15.Visible = False
    Text16.Visible = False
    Text17.Visible = False
    Text18.Visible = False
    Text19.Visible = False
    Text20.Visible = False
    Text21.Visible = False
    Text22.Visible = False
    Text23.Visible = False
    Text24.Visible = False
    Text25.Visible = False
    Text26.Visible = False
    Text27.Visible = False
    Text28.Visible = False
    Text46.Visible = False
    Text45.Visible = False
    Text44.Visible = False
End Sub

Sub Activar()

    Frame_total.Visible = True
    Check1.Visible = True
    Frame1.Visible = True
    Command1.Visible = True

    Label1.Visible = True
    Label2.Visible = True
    Label3.Visible = True
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    Label7.Visible = True
    Label8.Visible = True
    Label9.Visible = True
    Label10.Visible = True
    Label11.Visible = True
    Label12.Visible = True
    Label14.Visible = True
    Label15.Visible = True

    Text5.Visible = True
    Text6.Visible = True
    Text7.Visible = True
    Text8.Visible = True
    Text9.Visible = True
    Text10.Visible = True
    Text11.Visible = True
    Text46.Visible = True

    txt_iva1.Visible = True
    txt_iva2.Visible = True
    txt_iva3.Visible = True
    txt_iva4.Visible = True
    txt_iva5.Visible = True
    txt_iva6.Visible = True
    txt_iva7.Visible = True
    txt_iva8.Visible = True

    Text14.Visible = True
    Text15.Visible = True
    Text16.Visible = True
    Text17.Visible = True
    Text18.Visible = True
    Text19.Visible = True
    Text20.Visible = True
    Text21.Visible = True
    Text22.Visible = True
    Text23.Visible = True
    Text24.Visible = True
    Text25.Visible = True
    Text26.Visible = True
    Text27.Visible = True
    Text28.Visible = True
    Text44.Visible = True
    Text45.Visible = True
    Text46.Visible = True

End Sub

