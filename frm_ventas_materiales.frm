VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_ventas_materiales 
   Caption         =   "SISTEMA DE GESTIÓN Y CONTROL JER-GAS, C.A."
   ClientHeight    =   5400
   ClientLeft      =   2565
   ClientTop       =   945
   ClientWidth     =   8265
   Icon            =   "frm_ventas_materiales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   8265
   Begin VB.TextBox txt_objeto 
      Height          =   285
      Left            =   6360
      TabIndex        =   43
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Text8 
      DataField       =   "codigo"
      DataSource      =   "facturando"
      Height          =   285
      Left            =   600
      TabIndex        =   42
      Text            =   "Text8"
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt_zona 
      DataField       =   "id_ruta"
      DataSource      =   "clientes"
      Height          =   285
      Left            =   8160
      TabIndex        =   41
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txt_telefono_hab 
      DataField       =   "telefono_hab"
      DataSource      =   "clientes"
      Height          =   285
      Left            =   8160
      TabIndex        =   40
      Top             =   4800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txt_observaciones 
      DataField       =   "observaciones"
      DataSource      =   "clientes"
      Height          =   285
      Left            =   8160
      TabIndex        =   39
      Top             =   4440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txt_cedula 
      DataField       =   "cedula"
      DataSource      =   "clientes"
      Height          =   285
      Left            =   8160
      TabIndex        =   38
      Top             =   4080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txt_num_factura 
      DataField       =   "num_factura"
      DataSource      =   "pedidos"
      Height          =   285
      Left            =   8160
      TabIndex        =   37
      Top             =   3720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txt_num_control 
      DataField       =   "num_control"
      DataSource      =   "pedidos"
      Height          =   285
      Left            =   8160
      TabIndex        =   36
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      DataField       =   "id_pedido"
      DataSource      =   "relacion"
      Height          =   285
      Left            =   5160
      TabIndex        =   35
      Text            =   "Text7"
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text6 
      DataField       =   "codigo"
      DataSource      =   "resumen_inv2"
      Height          =   285
      Left            =   5760
      TabIndex        =   34
      Text            =   "Text6"
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text4 
      DataField       =   "codigo"
      DataSource      =   "resumen_inv1"
      Height          =   285
      Left            =   5760
      TabIndex        =   33
      Text            =   "Text4"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt_usuario 
      Height          =   285
      Left            =   3000
      TabIndex        =   32
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt_inst 
      DataField       =   "id_inst"
      DataSource      =   "pedidos"
      Height          =   285
      Left            =   3480
      TabIndex        =   31
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txt_fechando 
      DataField       =   "fecha_pedido"
      DataSource      =   "pedidos"
      Height          =   285
      Left            =   4080
      TabIndex        =   30
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      DataField       =   "codigo"
      DataSource      =   "clientes"
      Height          =   285
      Left            =   2160
      TabIndex        =   29
      Text            =   "Text3"
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text2 
      DataField       =   "codigo"
      DataSource      =   "materiales"
      Height          =   285
      Left            =   3000
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt_precio 
      DataField       =   "precio"
      DataSource      =   "materiales"
      Height          =   285
      Left            =   3720
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      DataField       =   "id_pedido"
      DataSource      =   "pedidos"
      Height          =   285
      Left            =   2160
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text5 
      DataField       =   "id_pedido"
      DataSource      =   "ventas"
      Height          =   285
      Left            =   2160
      TabIndex        =   24
      Text            =   "Text5"
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txt_codigo_mat 
      DataField       =   "codigo"
      DataSource      =   "materiales"
      Height          =   285
      Left            =   2160
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame6 
      Caption         =   "Nº Orden"
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
      TabIndex        =   21
      ToolTipText     =   "Suministre el Código del Cliente"
      Top             =   1560
      Width           =   2055
      Begin VB.TextBox txt_idpedido 
         Alignment       =   2  'Center
         DataField       =   "id_pedido"
         DataSource      =   "pedidos"
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
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   1815
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
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Suministre el Código para su ingreso."
      Top             =   3120
      Width           =   7935
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
         TabIndex        =   16
         ToolTipText     =   "Cantidad Actual"
         Top             =   360
         Width           =   2775
         Begin MSDataListLib.DataCombo txt_descripcion 
            Bindings        =   "frm_ventas_materiales.frx":08CA
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
      Begin VB.Frame Frame8 
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
         Left            =   3960
         TabIndex        =   28
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
         TabIndex        =   19
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
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "P/V Bs. F."
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
         TabIndex        =   17
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
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   240
            Width           =   1215
         End
      End
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
         TabIndex        =   15
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
   End
   Begin VB.Frame Frame17 
      Height          =   975
      Left            =   1920
      TabIndex        =   13
      Top             =   4320
      Width           =   4455
      Begin VB.CommandButton cmdsalir 
         Caption         =   "C&errar"
         Height          =   615
         Left            =   3000
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Cerrar"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdcancelar 
         Caption         =   "&Cancelar"
         Height          =   615
         Left            =   1560
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdprocesar 
         Caption         =   "&Procesar"
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1335
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
      TabIndex        =   10
      ToolTipText     =   "Suministre la Dirección del Cliente."
      Top             =   2280
      Width           =   7935
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
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   7695
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
      Left            =   3840
      TabIndex        =   8
      ToolTipText     =   "Suministre Apellido y Nombre del Cliente"
      Top             =   1560
      Width           =   4215
      Begin VB.TextBox txt_nombre 
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
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame Frame2 
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
      Left            =   2280
      TabIndex        =   0
      ToolTipText     =   "Suministre el Código del Cliente"
      Top             =   1560
      Width           =   1455
      Begin VB.TextBox txt_cliente 
         Alignment       =   2  'Center
         DataField       =   "codigo"
         DataSource      =   "pedidos"
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
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
      End
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
   Begin MSAdodcLib.Adodc pedidos 
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
   Begin MSAdodcLib.Adodc clientes 
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
   Begin MSAdodcLib.Adodc resumen_inv1 
      Height          =   330
      Left            =   6120
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
      Left            =   6120
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
   Begin MSAdodcLib.Adodc relacion 
      Height          =   375
      Left            =   3120
      Top             =   600
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
      RecordSource    =   "tbl_relacion_diaria"
      Caption         =   "relacion"
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
      Top             =   4440
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
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   -120
      Top             =   1020
      Width           =   15705
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
      Left            =   3600
      TabIndex        =   12
      Top             =   240
      Width           =   5055
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000B&
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   -120
      Top             =   960
      Width           =   15105
   End
End
Attribute VB_Name = "frm_ventas_materiales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdprocesar_Click()
 Dim Total1
 Dim marca
 
 With ventas.Recordset
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
        .MoveLast
    End If
        .AddNew
    !usuario_liq = Usuario
    !id_pedido = Me.txt_idpedido.Text
    !codigo = Me.txt_cliente.Text
    !status = "VI"
    !fecha_pedido = Me.txt_fechando.Text
    !descripcion = Me.txt_descripcion.Text
    !id_inst = Me.txt_inst.Text
    
    !monto_total = CCur(txt_precio_uni.Text)
    !cant_pedido = CInt(Me.txt_cant.Text)
    !iva = CCur(Me.txt_iva.Text)
    !precio = CCur(txt_monto.Text)
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

   With relacion.Recordset

    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
        .MoveLast
    End If
     .AddNew
       !usuario_liq = Usuario
       !id_pedido = txt_idpedido.Text
       !codigo = Me.txt_cliente.Text
       !descripcion = txt_descripcion.Text
       !status = "VI"
       !fecha_pedido = Me.txt_fechando.Text
       !id_inst = Me.txt_inst.Text
       !precio = CCur(txt_precio_uni.Text)
       !cant_pedido = CInt(Me.txt_cant.Text)
       !iva = CCur(Me.txt_iva.Text)
       !monto = CCur(txt_monto.Text)
       !marca = "2"
    .Update
   End With
   
   With facturando.Recordset

    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
        .MoveLast
    End If
    .AddNew
    
    !usuario_liq = Usuario
    !id_pedido = Me.txt_idpedido.Text
    !num_control = Me.txt_num_control.Text
    !num_factura = Me.txt_num_factura.Text
    
    !codigo = txt_cliente.Text
    !fecha_pedido = Me.txt_fechando.Text
    !cliente = Me.txt_nombre.Text
    !cedula = Me.txt_cedula.Text
    !direccion = Me.txt_direccion.Text
    !telefono_hab = Me.txt_telefono_hab.Text
    !observaciones = Me.txt_observaciones.Text
    !id_inst = Me.txt_inst.Text
    !descripcion = Me.txt_descripcion.Text
    !status = "VI"
    !id_ruta = Me.txt_zona.Text
    !marca = "2"
    !cant_pedido = CCur(Me.txt_cant.Text)
    !monto_fac = CCur(Me.txt_precio_uni.Text)
    !iva = CCur(Me.txt_iva.Text)
    !total_fac = CCur(Me.txt_monto.Text)

    .Update
End With
    
         MsgBox "La Venta Ha Sido Registrada, Presione Aceptar Para Finalizar ", vbInformation, "JerGas"
Unload Me
   
   

txt_descripcion.Text = ""
txt_precio_uni.Text = ""
txt_cant.Text = ""
txt_iva.Text = ""
txt_monto.Text = ""
enlace1 = ""
enlace2 = ""

     MsgBox "Proceso Finalizado", vbInformation, "JerGas"
Unload Me
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
txt_descripcion.Text = ""
txt_precio_uni.Text = ""
txt_cant.Text = ""
txt_iva.Text = ""
txt_monto.Text = ""

   With pedidos.Recordset
        pedidos.Recordset.MoveFirst
           strquery = "id_pedido = '" & enlace1 & "'"
        pedidos.Recordset.Find strquery
        
    !id_pedido = Me.txt_idpedido.Text
    !num_control = Me.txt_num_control.Text
    !num_factura = Me.txt_num_factura.Text
    !codigo = Me.txt_cliente.Text
    !fecha_pedido = Me.txt_fechando.Text
    !id_inst = Me.txt_inst.Text
    
   With clientes.Recordset
        clientes.Recordset.MoveFirst
           strquery2 = "codigo = '" & enlace2 & "'"
        clientes.Recordset.Find strquery2
        !cliente = Me.txt_nombre.Text
        !direccion = Me.txt_direccion.Text
        !cedula = Me.txt_cedula.Text
        !telefono_hab = Me.txt_telefono_hab.Text
        !observaciones = Me.txt_observaciones.Text
        !id_ruta = Me.txt_zona.Text
     End With
   End With
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

Private Sub txt_cant_KeyPress(KeyAscii As Integer)
 
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
   
   End Sub

Private Sub txt_precio_uni_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then SendKeys "{tab}"
 '   If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
 '   KeyAscii = Asc(UCase(Chr(KeyAscii)))
 '       If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
   
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

