VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm_ventas_materiales2 
   Caption         =   "SISTEMA DE GESTIÓN Y CONTROL JER-GAS, C.A."
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10200
   Icon            =   "frm_ventas_materiales2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   10200
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      DataField       =   "precio_venta"
      DataSource      =   "materiales"
      Height          =   375
      Left            =   7440
      TabIndex        =   44
      Text            =   "Text5"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      DataField       =   "id_inst"
      DataSource      =   "instalacion"
      Height          =   285
      Left            =   2040
      TabIndex        =   43
      Text            =   "Text3"
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt_ruta 
      DataField       =   "id_ruta"
      DataSource      =   "clientes"
      Height          =   285
      Left            =   240
      TabIndex        =   42
      Top             =   6960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txt_telefono_hab 
      DataField       =   "telefono_hab"
      DataSource      =   "clientes"
      Height          =   285
      Left            =   240
      TabIndex        =   41
      Top             =   6600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt_observaciones 
      DataField       =   "observaciones"
      DataSource      =   "clientes"
      Height          =   285
      Left            =   240
      TabIndex        =   40
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      DataField       =   "id_pedido"
      DataSource      =   "facturando"
      Height          =   285
      Left            =   2040
      TabIndex        =   39
      Text            =   "Text2"
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      DataField       =   "id_pedido"
      DataSource      =   "ventas"
      Height          =   285
      Left            =   2040
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text4 
      DataField       =   "codigo"
      DataSource      =   "resumen_inv1"
      Height          =   285
      Left            =   7320
      TabIndex        =   35
      Text            =   "Text4"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text6 
      DataField       =   "codigo"
      DataSource      =   "resumen_inv2"
      Height          =   285
      Left            =   7320
      TabIndex        =   34
      Text            =   "Text6"
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt_codigo_mat 
      DataField       =   "codigo"
      DataSource      =   "materiales"
      Height          =   285
      Left            =   2040
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txt_date 
      Height          =   285
      Left            =   8520
      TabIndex        =   32
      Text            =   "txt_date"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "DATOS DEL CLIENTE"
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
      Height          =   2895
      Left            =   840
      TabIndex        =   19
      Top             =   1920
      Width           =   7935
      Begin VB.Frame Frame10 
         Caption         =   "T/Inst."
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
         Left            =   6600
         TabIndex        =   37
         ToolTipText     =   "Suministre el Código del Cliente"
         Top             =   360
         Width           =   975
         Begin VB.TextBox txt_inst 
            Alignment       =   2  'Center
            DataField       =   "id_inst"
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
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Código "
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
         Left            =   120
         TabIndex        =   28
         ToolTipText     =   "Suministre el Código del Cliente"
         Top             =   360
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
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   240
            Width           =   1575
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
         Left            =   120
         TabIndex        =   26
         ToolTipText     =   "Suministre Apellido y Nombre del Cliente"
         Top             =   1200
         Width           =   6135
         Begin VB.TextBox txt_clientes 
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
            TabIndex        =   27
            Top             =   240
            Width           =   5895
         End
      End
      Begin VB.Frame Frame12 
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
         TabIndex        =   24
         ToolTipText     =   "Suministre el Nº de Contrato del Cliente"
         Top             =   360
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
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame6 
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
         Left            =   4200
         TabIndex        =   22
         ToolTipText     =   "Suministre la Cédula del Cliente."
         Top             =   360
         Width           =   2055
         Begin VB.TextBox txt_cedula 
            Alignment       =   2  'Center
            DataField       =   "cedula"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0"
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
            TabIndex        =   23
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
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Suministre la Dirección del Cliente."
         Top             =   2040
         Width           =   7695
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
            TabIndex        =   21
            Top             =   240
            Width           =   7455
         End
      End
   End
   Begin VB.Frame Frame17 
      Height          =   975
      Left            =   2640
      TabIndex        =   18
      Top             =   6240
      Width           =   4455
      Begin VB.CommandButton cmd_procesar 
         Caption         =   "&Procesar"
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmd_cancelar 
         Caption         =   "&Cancelar"
         Height          =   615
         Left            =   1560
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "C&errar"
         Height          =   615
         Left            =   3000
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Cerrar"
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "INGRESE MATERIALES"
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
      Left            =   840
      TabIndex        =   10
      ToolTipText     =   "Suministre el Código para su ingreso."
      Top             =   4920
      Width           =   7935
      Begin VB.Frame Frame7 
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
         Left            =   3000
         TabIndex        =   17
         ToolTipText     =   "Cantidad Actual"
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
            TabIndex        =   3
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Total Bs. F."
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
         TabIndex        =   15
         ToolTipText     =   "Cantidad Actual"
         Top             =   360
         Width           =   1455
         Begin VB.TextBox txt_monto 
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
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
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
         Left            =   5280
         TabIndex        =   13
         ToolTipText     =   "Cantidad Actual"
         Top             =   360
         Width           =   975
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
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "P / Venta"
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
         Left            =   3960
         TabIndex        =   12
         ToolTipText     =   "Cantidad Actual"
         Top             =   360
         Width           =   1215
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
            Left            =   120
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
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
         TabIndex        =   11
         ToolTipText     =   "Cantidad Actual"
         Top             =   360
         Width           =   2775
         Begin MSDataListLib.DataCombo txt_descripcion 
            Bindings        =   "frm_ventas_materiales2.frx":08CA
            Height          =   315
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   2535
            _ExtentX        =   4471
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
   Begin VB.CommandButton Busquedad_avanzadas 
      Caption         =   "Búsqueda"
      Height          =   375
      Index           =   0
      Left            =   8640
      TabIndex        =   9
      Tag             =   "Lista todos los inmuebles registrados"
      Top             =   1080
      Width           =   1095
   End
   Begin MSDataListLib.DataCombo Dcmb_Buscar 
      Bindings        =   "frm_ventas_materiales2.frx":08E3
      Height          =   315
      Left            =   3480
      TabIndex        =   8
      ToolTipText     =   "Pulse doble click para cambiar el tipo de busqueda, después de presionar búsqueda avanzada"
      Top             =   1080
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      ListField       =   "codigo"
      BoundColumn     =   "codigo"
      Text            =   ""
      Object.DataMember      =   ""
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
   Begin MSAdodcLib.Adodc ventas 
      Height          =   375
      Left            =   0
      Top             =   360
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
      RecordSource    =   "tbl_resumen_ventas_materiales"
      Caption         =   "ventas"
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
   Begin MSAdodcLib.Adodc clientes 
      Height          =   375
      Left            =   0
      Top             =   720
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
   Begin MSAdodcLib.Adodc inventario 
      Height          =   375
      Left            =   3360
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc materiales 
      Height          =   375
      Left            =   0
      Top             =   0
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
   Begin MSAdodcLib.Adodc resumen_inv1 
      Height          =   330
      Left            =   7680
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
      Left            =   7680
      Top             =   240
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
   Begin MSAdodcLib.Adodc facturando 
      Height          =   375
      Left            =   0
      Top             =   1080
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
   Begin MSAdodcLib.Adodc instalacion 
      Height          =   375
      Left            =   0
      Top             =   1440
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
      RecordSource    =   "tbl_instalacion"
      Caption         =   "instalacion"
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
   Begin MSForms.TextBox txt_fecha_entrega 
      Bindings        =   "frm_ventas_materiales2.frx":08FA
      CausesValidation=   0   'False
      Height          =   300
      Left            =   5280
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   1455
      VariousPropertyBits=   746604575
      BackColor       =   -2147483633
      Size            =   "2566;529"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txt_fecha_pedido 
      Bindings        =   "frm_ventas_materiales2.frx":0915
      CausesValidation=   0   'False
      Height          =   300
      Left            =   3600
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   1455
      VariousPropertyBits=   746604575
      BackColor       =   -2147483633
      Size            =   "2566;529"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label_titulo 
      BackColor       =   &H80000001&
      Caption         =   "  VENTAS DE MATERIALES"
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
      Left            =   3840
      TabIndex        =   7
      Top             =   240
      Width           =   7695
   End
   Begin VB.Label LABEL_BUSCA 
      BackStyle       =   0  'Transparent
      Caption         =   "Búsqueda por Código: "
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
      TabIndex        =   0
      Top             =   1200
      Width           =   3615
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
      Width           =   11505
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
Attribute VB_Name = "frm_ventas_materiales2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Total1
Dim marca

Private Sub Busquedad_avanzadas_Click(Index As Integer)
            Busq_Avanzada = True
            clientes.CommandType = adCmdText
            clientes.RecordSource = "select * from tbl_clientes WHERE codigo <> '' ORDER BY codigo ASC"
            clientes.Refresh
            Call Dcmb_Buscar_Click(1)
End Sub

Private Sub buscar_codigo()
On Error GoTo ControlError
Dim strquery, RESP

If Not Busq_Avanzada And ((Dcmb_Buscar.Text Like "%*%" Or Dcmb_Buscar.Text Like "%*" Or Dcmb_Buscar.Text Like "*%") Or (Me.clientes.Recordset.RecordCount = 0)) Then
    
    Me.clientes.CommandType = adCmdText
    Me.clientes.RecordSource = "SELECT * FROM tbl_clientes WHERE tbl_clientes.codigo like '" & Dcmb_Buscar.Text & "' ORDER BY codigo"
    Me.clientes.Refresh

    If clientes.Recordset.EOF Then
        MsgBox "Código suministrado no encontrado, por favor verifique ", vbInformation, "JerGas C.A."
        Dcmb_Buscar.Text = ""
        Dcmb_Buscar.SetFocus
     Else
        If Me.clientes.Recordset.RecordCount > 1 Then
            MsgBox Me.clientes.Recordset.RecordCount & " encontrados"
            Busq_Avanzada = True
        End If
        
    End If
    
Else
    
    clientes.Recordset.MoveFirst
    strquery = "codigo = '" & Dcmb_Buscar.Text & "'"
    clientes.Recordset.Find strquery
    
    If clientes.Recordset.EOF Then
            MsgBox "Código suministrado no encontrado, por favor verifique ", vbInformation, "SIAGEP"
            Dcmb_Buscar.Text = ""
            Dcmb_Buscar.SetFocus
    End If
    

End If

Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "JerGas")
        Case 3001
            v = MsgBox("Nombre suministrado no encontrado", vbOKOnly, "JerGas")
    End Select

End Sub

Private Sub buscar_pedidos()
On Error GoTo ControlError
hist_pedidos.CommandType = adCmdText
hist_pedidos.RecordSource = "SELECT * FROM tbl_pedidos WHERE tbl_pedidos.codigo = '" & Me.txt_clientes.Text & "' ORDER BY fecha_pedido DESC"
hist_pedidos.Refresh

Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "JerGas")
        Case 3001
            v = MsgBox("Nombre suministrado no encontrado", vbOKOnly, "JerGas")
    End Select
    
End Sub

Private Sub cmd_eliminar_Click()
Dim resta
On Error GoTo control_error

'Desabilita el botón de aceptar
'Me.cmd_eliminar.Enabled = False

Screen.MousePointer = 11

If DGrid_pedidos.SelBookmarks.Count = 0 Then
    
    MsgBox "No se hallaron Pedidos marcados para Eliminar."
'    Me.cmd_eliminar.Enabled = True
    Screen.MousePointer = 0
    Exit Sub

End If
    pedidos.Recordset.MoveFirst
    strquery = "id_pedido = '" & DGrid_pedidos.Columns(0).Text & "'"
    pedidos.Recordset.Find strquery
       If pedidos.Recordset.EOF Then
                MsgBox "Nºde Planilla suministrada no encontrada, por favor verifique ", vbInformation, "JerGas C.A."
          Screen.MousePointer = 0
       Else
          pedidos.Recordset.Delete
       FGNRO_LIQ_RESTA
       End If
    
hist_pedidos.Refresh
pedidos.Refresh
Screen.MousePointer = 0

'With control.Recordset
'    !Nro_liquida_ult = Me.txt_resta.Text
'      resta = Me.txt_resta.Text - 1
'    !Nro_liquida_ult = resta
'    .Update
'End With

Exit Sub

control_error:
Screen.MousePointer = 0
    MsgBox Err.Description

End Sub

Private Sub Dcmb_Buscar_Click(Area As Integer)
If Area = 2 Then
        If Dcmb_Buscar.ListField = "codigo" Then
            If Dcmb_Buscar.Text <> "" Then
                
                Call buscar_codigo
                Call buscar_pedidos
            Else
                Exit Sub
            End If
        End If
        
        If Dcmb_Buscar.ListField = "contrato_num" Then
            If Dcmb_Buscar.Text <> "" Then
                Call buscar_contrato_num
                Call buscar_pedidos
            Else
                Exit Sub
            End If
        End If

        If Dcmb_Buscar.ListField = "cliente" Then
            If Dcmb_Buscar.Text <> "" Then
                Call buscar_cliente
                Call buscar_pedidos
            Else
                Exit Sub
            End If
        End If

        If Dcmb_Buscar.ListField = "cedula" Then
            If Dcmb_Buscar.Text <> "" Then
                Call buscar_cedula
                Call buscar_pedidos
            Else
                Exit Sub
            End If
        End If
        If Dcmb_Buscar.ListField = "direccion" Then
            If Dcmb_Buscar.Text <> "" Then
                Call buscar_direccion
                Call buscar_pedidos
            Else
                Exit Sub
            End If
        End If
End If
End Sub
Private Sub buscar_contrato_num()

On Error GoTo ControlError

Dim strquery, RESP

If Not Busq_Avanzada And ((Dcmb_Buscar.Text Like "%*%" Or Dcmb_Buscar.Text Like "%*" Or Dcmb_Buscar.Text Like "*%") Or (Me.clientes.Recordset.RecordCount = 0)) Then
    
    Me.clientes.CommandType = adCmdText
    
    Me.clientes.RecordSource = "SELECT * FROM tbl_clientes WHERE tbl_clientes.contrato_num like '" & Dcmb_Buscar.Text & "' ORDER BY contrato_num"
    
    Me.clientes.Refresh

    If clientes.Recordset.EOF Then
    
        MsgBox "Nº de Contrato suministrado no encontrado, por favor verifique ", vbInformation, "JerGas C.A."
        
        Dcmb_Buscar.Text = ""
        
        Dcmb_Buscar.SetFocus
        
  
    Else
    
        If Me.clientes.Recordset.RecordCount > 1 Then
        
            MsgBox Me.clientes.Recordset.RecordCount & " encontrados"
            
            Busq_Avanzada = True
            
        End If
        
        
    End If
    
Else
    
    clientes.Recordset.MoveFirst
    
    strquery = "contrato_num = '" & Dcmb_Buscar.Text & "'"

    clientes.Recordset.Find strquery
    
    If clientes.Recordset.EOF Then
    
            MsgBox "Nº de Contrato suministrado no encontrado, por favor verifique ", vbInformation, "SIAGEP"
            
            Dcmb_Buscar.Text = ""
            
            Dcmb_Buscar.SetFocus
            
                    
    Else
    
        
    End If
    

End If

Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "JerGas")
        Case 3001
            v = MsgBox("Contrato suministrado no encontrado", vbOKOnly, "JerGas")
    End Select

End Sub
Private Sub buscar_cedula()

On Error GoTo ControlError

Dim strquery, RESP

If Not Busq_Avanzada And ((Dcmb_Buscar.Text Like "%*%" Or Dcmb_Buscar.Text Like "%*" Or Dcmb_Buscar.Text Like "*%") Or (Me.clientes.Recordset.RecordCount = 0)) Then
    
    Me.clientes.CommandType = adCmdText
    
    Me.clientes.RecordSource = "SELECT * FROM tbl_clientes WHERE tbl_clientes.cedula like '" & Dcmb_Buscar.Text & "' ORDER BY cedula"
    
    Me.clientes.Refresh

    If clientes.Recordset.EOF Then
    
        MsgBox "Cédula del cliente suministrado no encontrado, por favor verifique ", vbInformation, "JerGas C.A."
        
        Dcmb_Buscar.Text = ""
        
        Dcmb_Buscar.SetFocus
        
        
    Else
    
        If Me.clientes.Recordset.RecordCount > 1 Then
        
            MsgBox Me.clientes.Recordset.RecordCount & " encontrados"
            
            Busq_Avanzada = True
            
        End If
        
        
    End If
    
Else
    
    clientes.Recordset.MoveFirst
    
    strquery = "cedula = '" & Dcmb_Buscar.Text & "'"

    clientes.Recordset.Find strquery
    
    If clientes.Recordset.EOF Then
    
            MsgBox "Cédula del Cliente suministrado no encontrado, por favor verifique ", vbInformation, "SIAGEP"
            
            Dcmb_Buscar.Text = ""
            
            Dcmb_Buscar.SetFocus
            
 '           Call habilitar_botones(False)
                    
 '   Else
    
'            Call habilitar_botones(True)
        
    End If
    

End If

Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "JerGas")
        Case 3001
            v = MsgBox("Cédula del Cliente suministrado no encontrado", vbOKOnly, "JerGas")
    End Select

End Sub

Private Sub buscar_cliente()

On Error GoTo ControlError

Dim strquery, RESP

If Not Busq_Avanzada And ((Dcmb_Buscar.Text Like "%*%" Or Dcmb_Buscar.Text Like "%*" Or Dcmb_Buscar.Text Like "*%") Or (Me.clientes.Recordset.RecordCount = 0)) Then
    
    Me.clientes.CommandType = adCmdText
    
    Me.clientes.RecordSource = "SELECT * FROM tbl_clientes WHERE tbl_clientes.cliente like '" & Dcmb_Buscar.Text & "' ORDER BY cliente"
    
    Me.clientes.Refresh

    If clientes.Recordset.EOF Then
    
        MsgBox "Nombre del cliente suministrado no encontrado, por favor verifique ", vbInformation, "JerGas C.A."
        
        Dcmb_Buscar.Text = ""
        
        Dcmb_Buscar.SetFocus
        
       
    Else
    
        If Me.clientes.Recordset.RecordCount > 1 Then
        
            MsgBox Me.clientes.Recordset.RecordCount & " encontrados"
            
            Busq_Avanzada = True
            
        End If
        
        
    End If
    
Else
    
    clientes.Recordset.MoveFirst
    
    strquery = "cliente = '" & Dcmb_Buscar.Text & "'"

    clientes.Recordset.Find strquery
    
    If clientes.Recordset.EOF Then
    
            MsgBox "Nombre del Cliente suministrado no encontrado, por favor verifique ", vbInformation, "JerGas"
            
            Dcmb_Buscar.Text = ""
            
            Dcmb_Buscar.SetFocus
            
                       
    Else
    
             
    End If
    

End If

Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "JerGas")
        Case 3001
            v = MsgBox("Nombre del Cliente suministrado no encontrado", vbOKOnly, "JerGas")
    End Select

End Sub

Private Sub buscar_direccion()

On Error GoTo ControlError

Dim strquery, RESP

If Not Busq_Avanzada And ((Dcmb_Buscar.Text Like "%*%" Or Dcmb_Buscar.Text Like "%*" Or Dcmb_Buscar.Text Like "*%") Or (Me.clientes.Recordset.RecordCount = 0)) Then
    
    Me.clientes.CommandType = adCmdText
    
    Me.clientes.RecordSource = "SELECT * FROM tbl_clientes WHERE tbl_clientes.direccion like '" & Dcmb_Buscar.Text & "' ORDER BY direccion"
    
    Me.clientes.Refresh

    If clientes.Recordset.EOF Then
    
        MsgBox "Direccion del cliente suministrada no encontrada, por favor verifique ", vbInformation, "JerGas C.A."
        
        Dcmb_Buscar.Text = ""
        
        Dcmb_Buscar.SetFocus
        
        
    Else
    
        If Me.clientes.Recordset.RecordCount > 1 Then
        
            MsgBox Me.clientes.Recordset.RecordCount & " encontrados"
            
            Busq_Avanzada = True
            
        End If
        
        
    End If
    
Else
    
    clientes.Recordset.MoveFirst
    
    strquery = "direccion = '" & Dcmb_Buscar.Text & "'"

    clientes.Recordset.Find strquery
    
    If clientes.Recordset.EOF Then
    
            MsgBox "Dirección del Cliente suministrado no encontrado, por favor verifique ", vbInformation, "JerGas"
            
            Dcmb_Buscar.Text = ""
            
            Dcmb_Buscar.SetFocus
            
                    
    Else
    
        
    End If
    

End If

Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "JerGas")
        Case 3001
            v = MsgBox("Dirección del Cliente suministrado no encontrado", vbOKOnly, "JerGas")
    End Select

End Sub

Private Sub Dcmb_Buscar_DblClick(Area As Integer)
'Esta funci[on redefine el tipo de busqueda
'If Busq_Avanzada Then

    Me.Dcmb_Buscar.ToolTipText = "Doble click para alternar el tipo de busqueda"
    
    If Dcmb_Buscar.ListField = "codigo" Then
        'Si es bif pasa a ape
        If Busq_Avanzada Then
        
            Me.clientes.CommandType = adCmdText

            Me.clientes.RecordSource = "select * from tbl_clientes WHERE tbl_clientes.contrato_num <> '' ORDER BY contrato_num ASC"

            Me.clientes.Refresh
            
        End If

        Dcmb_Buscar.ListField = "contrato_num"

        Dcmb_Buscar.Text = ""

        LABEL_BUSCA.Caption = "Búsqueda por Nº de Contrato:"

        Exit Sub
        
    End If

    If Dcmb_Buscar.ListField = "contrato_num" Then
    
        'Si es ape pasa a cod cata
        If Busq_Avanzada Then
        
            Me.clientes.CommandType = adCmdText

            Me.clientes.RecordSource = "select * from tbl_clientes WHERE tbl_clientes.cliente <> '' ORDER BY cliente ASC"

            Me.clientes.Refresh
            
        End If
        
        Dcmb_Buscar.ListField = "cliente"

        Dcmb_Buscar.Text = ""

        LABEL_BUSCA.Caption = "Búsqueda por Cliente:"

        Exit Sub
        
    End If

    If Dcmb_Buscar.ListField = "cliente" Then

        'Si es cod pasa a cedula
        If Busq_Avanzada Then
        
            clientes.CommandType = adCmdText

            clientes.RecordSource = "select * from tbl_clientes WHERE tbl_clientes.cedula <> '' ORDER BY cedula ASC"

            clientes.Refresh
            
        End If
        
        Dcmb_Buscar.ListField = "cedula"

        Dcmb_Buscar.Text = ""

        LABEL_BUSCA.Caption = "Búsqueda por Cédula: "

        Exit Sub

    End If

    If Dcmb_Buscar.ListField = "cedula" Then
    
        'Si es cedual pasa a direccion
        If Busq_Avanzada Then
        
            clientes.CommandType = adCmdText

            clientes.RecordSource = "select * from tbl_clientes WHERE tbl_clientes.direccion <> '' ORDER BY direccion ASC"

            clientes.Refresh
            
        End If
        
        Dcmb_Buscar.ListField = "direccion"

        Dcmb_Buscar.Text = ""

        LABEL_BUSCA.Caption = "Búsqueda por Dirección: "

        Exit Sub
        
    End If

    If Dcmb_Buscar.ListField = "direccion" Then

        'Si es direccion pasa a bif
        If Busq_Avanzada Then
        
            clientes.CommandType = adCmdText

            clientes.RecordSource = "select * from tbl_clientes WHERE codigo <> '' ORDER BY codigo ASC"

            clientes.Refresh
            
        End If
        
        Dcmb_Buscar.ListField = "codigo"

        Dcmb_Buscar.Text = ""

        LABEL_BUSCA.Caption = "Búsqueda por Código: "

        Exit Sub
    End If
End Sub

Private Sub Form_Load()
Call actualizar_cn("SQL Server")
Me.txt_fecha_pedido = Date
Me.txt_fecha_entrega = DateAdd("d", 1, Date)
Me.txt_date = Date

txt_codigo.Text = ""
txt_contrato.Text = ""
txt_cedula.Text = ""
txt_inst.Text = ""
txt_clientes.Text = ""
txt_direccion.Text = ""
txt_descripcion.Text = ""
txt_precio_uni.Text = ""
txt_cant.Text = ""
txt_iva.Text = ""
txt_monto.Text = ""
End Sub

Private Sub cmd_cerrar_Click()
Unload Me
End Sub

Private Sub cmd_procesar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_procesar.FontBold = True
Me.cmd_cancelar.FontBold = False
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = True
Me.cmd_procesar.FontBold = False
Me.cmd_cancelar.FontBold = False
End Sub

Private Sub cmd_cancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_procesar.FontBold = False
Me.cmd_cancelar.FontBold = True
End Sub

Private Sub cmd_procesar_Click()

On Error GoTo ControlError

Gcod_planilla = FGNRO_LIQ()
Gcod_control = FGNRO_CONTROL()
Gcod_factura = FGNRO_FACTURA()
 
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
    !cliente = Me.txt_clientes.Text
    !cedula = Me.txt_cedula.Text
    !direccion = Me.txt_direccion.Text
    !telefono_hab = Me.txt_telefono_hab.Text
    !observaciones = Me.txt_observaciones.Text
    !id_inst = Me.txt_inst.Text
    !descripcion = Me.txt_descripcion.Text
    !status = "VI"
    !id_ruta = Me.txt_ruta.Text
    !marca = "2"
    !cant_pedido = Me.txt_cant.Text
    !monto_fac = CCur(Me.txt_precio_uni.Text)
    !iva = CCur(Me.txt_iva.Text)
    !total_fac = CCur(Me.txt_monto.Text)

    .Update
End With

With ventas.Recordset
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
        .MoveLast
    End If
        .AddNew
    !usuario_liq = Usuario
    !id_pedido = Gcod_planilla
    !num_control = Gcod_control
    !num_factura = Gcod_factura
    
    !codigo = Me.txt_codigo.Text
    !status = "CA"
    !fecha_pedido = Me.txt_fecha_pedido.Text
    !fecha_cancel = Me.txt_fecha_pedido.Text
    !descripcion = Me.txt_descripcion.Text
    !id_inst = Me.txt_inst.Text
    !monto_total = CCur(Me.txt_monto.Text)
    !cant_pedido = CInt(Me.txt_cant.Text)
    !iva = CCur(Me.txt_iva.Text)
    !precio = CCur(txt_precio_uni.Text)
      .Update
    End With

    With materiales.Recordset
       mvBookMark = .Bookmark
       Me.materiales.RecordSource = "SELECT * FROM tbl_materiales WHERE tbl_materiales.codigo = '" & Me.txt_codigo_mat.Text & "' ORDER BY codigo"
       !cant_actual = !cant_actual - CInt(Me.txt_cant.Text)
       .Update
    End With

    If txt_codigo_mat.Text >= 1 And txt_codigo_mat.Text <= 4 Then
        marca = Me.txt_codigo_mat.Text
        
         With resumen_inv1.Recordset
           mvBookMark = .Bookmark
           .MoveFirst
           .Find "codigo =" & Me.txt_codigo_mat.Text & ""
              !cil_lleno = !cil_lleno - CInt(Me.txt_cant.Text)
              !cil_vacio = !cil_vacio + CInt(Me.txt_cant.Text)
           .Update
         End With
           Me.resumen_inv1.Refresh
     End If

     If txt_codigo_mat.Text >= 5 And txt_codigo_mat.Text <= 14 Then
        marca = Me.txt_codigo_mat.Text
        
         With resumen_inv2.Recordset
           mvBookMark = .Bookmark
           .MoveFirst
           .Find "codigo =" & Me.txt_codigo_mat.Text & ""
              !cant_actual = !cant_actual - CInt(Me.txt_cant.Text)
              !cant_inst = !cant_inst + CInt(Me.txt_cant.Text)
           .Update
         End With
           Me.resumen_inv2.Refresh
      End If


txt_codigo.Text = ""
txt_contrato.Text = ""
txt_cedula.Text = ""
txt_inst.Text = ""
txt_clientes.Text = ""
txt_direccion.Text = ""
txt_descripcion.Text = ""
txt_precio_uni.Text = ""
txt_cant.Text = ""
txt_iva.Text = ""
txt_monto.Text = ""

     MsgBox "Proceso Finalizado", vbInformation, "JerGas"
Unload Me


Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "JerGas")
    End Select
End Sub

Private Sub txt_descripcion_Click(Area As Integer)
Dim Total
On Error GoTo ControlError
If (Area = 2) Then

    strquery = "codigo = '" & txt_descripcion.BoundText & "'"
    materiales.Recordset.Find strquery
        If materiales.Recordset.EOF Then
            txt_descripcion.Text = ""
        End If
End If

Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "JerGas")
        End Select
End Sub

Private Sub txt_cant_LostFocus()

            If txt_cant.Text <> "" Then
             
             With materiales.Recordset
        materiales.Recordset.MoveFirst
           strquery = "descripcion = '" & Me.txt_descripcion.Text & "'"
        materiales.Recordset.Find strquery

            precio = !precio_venta
                 impuesto = !iva
                
                 cant_venta = Me.txt_cant.Text * precio
                      paso1 = CCur(cant_venta / 1.09)
                      paso2 = Round(paso1, 2)
                      
                 total_iva = CCur(paso2 * impuesto) / 100
                 paso3 = Round(total_iva, 2)
                   
                 precio_unitario = paso2

                 Me.txt_precio_uni = CCur(precio_unitario)
                 Me.txt_iva.Text = paso3
                 Me.txt_monto.Text = CCur(cant_venta)

                .Update
            End With
  End If
End Sub

Private Sub txt_cant_KeyPress(KeyAscii As Integer)
 
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
   
   End Sub

