VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm_pedidos 
   Caption         =   "SISTEMA DE GESTIÓN Y CONTROL JER-GAS, C.A."
   ClientHeight    =   9075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13455
   Icon            =   "frm_pedidos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9075
   ScaleWidth      =   13455
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_num_factura 
      Height          =   285
      Left            =   4800
      TabIndex        =   119
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txt_num_control 
      Height          =   285
      Left            =   3720
      TabIndex        =   118
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txt_control 
      Height          =   285
      Left            =   2520
      TabIndex        =   117
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      DataField       =   "id_pedido"
      DataSource      =   "facturando"
      Height          =   285
      Left            =   4920
      TabIndex        =   115
      Text            =   "Text6"
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text2 
      DataField       =   "id_pedido"
      DataSource      =   "relacion"
      Height          =   285
      Left            =   2040
      TabIndex        =   114
      Text            =   "Text2"
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSAdodcLib.Adodc relacion 
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
   Begin VB.TextBox txt_resta 
      DataField       =   "Nro_liquida_ult"
      DataSource      =   "control"
      Height          =   285
      Left            =   6600
      TabIndex        =   113
      Top             =   600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txt_id_inst 
      DataField       =   "id_inst"
      DataSource      =   "hist_pedidos"
      Height          =   375
      Left            =   5520
      TabIndex        =   111
      Top             =   8640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt_cant_pedido 
      DataField       =   "cant_pedido"
      DataSource      =   "hist_pedidos"
      Height          =   375
      Left            =   6120
      TabIndex        =   110
      Top             =   8640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text3 
      DataField       =   "cilindro"
      DataSource      =   "inventario"
      Height          =   375
      Left            =   6720
      TabIndex        =   109
      Top             =   8640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      DataField       =   "cant_actual"
      DataSource      =   "materiales"
      Height          =   375
      Left            =   7800
      TabIndex        =   108
      Top             =   8640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   8760
      TabIndex        =   107
      Top             =   8640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txt_codigo 
      DataField       =   "codigo"
      DataSource      =   "res_inventario"
      Height          =   375
      Left            =   5040
      TabIndex        =   106
      Top             =   8640
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSAdodcLib.Adodc ventas 
      Height          =   375
      Left            =   4200
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
      RecordSource    =   "tbl_resumen_venta"
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
   Begin VB.TextBox txt_pedidos_liq 
      DataField       =   "id_pedido"
      DataSource      =   "liquidado"
      Height          =   285
      Left            =   12000
      TabIndex        =   99
      Text            =   "txt_pedidos_liq"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txt_date 
      Height          =   285
      Left            =   12000
      TabIndex        =   98
      Text            =   "txt_date"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   0
      Top             =   1440
   End
   Begin VB.TextBox txt_id_ruta_chofer 
      DataField       =   "id_ruta"
      DataSource      =   "chofer"
      Height          =   285
      Left            =   13440
      TabIndex        =   87
      Text            =   "txt_id_ruta_chofer"
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txt_id_tbl_chofer 
      DataField       =   "id_chofer"
      DataSource      =   "chofer"
      Height          =   285
      Left            =   13440
      TabIndex        =   86
      Text            =   "txt_id_tbl_chofer"
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txt_cedula_chofer 
      DataField       =   "cedula"
      DataSource      =   "chofer"
      Height          =   285
      Left            =   13440
      TabIndex        =   85
      Text            =   "txt_cedula_chofer"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txt_nombre_chofer 
      DataField       =   "nombre"
      DataSource      =   "chofer"
      Height          =   285
      Left            =   13440
      TabIndex        =   84
      Text            =   "txt_nombre_chofer"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txt_id_ruta 
      DataField       =   "id_ruta"
      DataSource      =   "despacho"
      Height          =   285
      Left            =   12000
      TabIndex        =   83
      Text            =   "txt_id_ruta"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txt_des_fecha_pedido 
      DataField       =   "fecha_pedido"
      DataSource      =   "despacho"
      Height          =   285
      Left            =   12000
      TabIndex        =   82
      Text            =   "txt_des_fecha_pedido"
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txt_id_chofer 
      DataField       =   "id_chofer"
      DataSource      =   "despacho"
      Height          =   285
      Left            =   12000
      TabIndex        =   81
      Text            =   "txt_id_chofer"
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txt_id_pedidos 
      DataField       =   "id_pedido"
      DataSource      =   "pedidos"
      Height          =   285
      Left            =   13440
      TabIndex        =   80
      Text            =   "id_pedidos"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc clientes 
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   7455
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   11535
      Begin VB.CommandButton cmd_ventas_materiales 
         Caption         =   "Ventas de Materiales"
         Enabled         =   0   'False
         Height          =   735
         Left            =   2040
         TabIndex        =   112
         TabStop         =   0   'False
         Top             =   6240
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         DataField       =   "fecha_venta"
         DataSource      =   "resumen_mensual_ventas"
         Height          =   375
         Left            =   0
         TabIndex        =   105
         Top             =   7080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txt_inst 
         Height          =   375
         Left            =   1200
         TabIndex        =   104
         Top             =   6840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txt_cant 
         Height          =   375
         Left            =   1920
         TabIndex        =   103
         Top             =   6840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txt_fecha 
         Height          =   375
         Left            =   0
         TabIndex        =   102
         Top             =   6840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txt_monto 
         Height          =   375
         Left            =   2640
         TabIndex        =   101
         Top             =   6840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmd_eliminar 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Height          =   735
         Left            =   3480
         TabIndex        =   100
         Top             =   6240
         Width           =   1455
      End
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "Cerrar"
         Height          =   735
         Left            =   7800
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   6240
         Width           =   1455
      End
      Begin VB.CommandButton cmd_liquidacion 
         Caption         =   "Liquidación"
         Enabled         =   0   'False
         Height          =   735
         Left            =   600
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   6240
         Width           =   1455
      End
      Begin VB.CommandButton cmd_clientes 
         Caption         =   "Editar Clientes"
         Enabled         =   0   'False
         Height          =   735
         Left            =   6360
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   6240
         Width           =   1455
      End
      Begin VB.CommandButton cmd_estado 
         Caption         =   "Estado de Cuenta"
         Enabled         =   0   'False
         Height          =   735
         Left            =   4920
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   6240
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Caption         =   "DATOS DEL CLIENTE"
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
         Height          =   3735
         Left            =   0
         TabIndex        =   51
         Top             =   120
         Width           =   5175
         Begin VB.TextBox txt_observaciones 
            DataField       =   "observaciones"
            DataSource      =   "clientes"
            Height          =   285
            Left            =   3120
            TabIndex        =   116
            Top             =   1800
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label Label2 
            Caption         =   "Nº de Contrato"
            Height          =   240
            Left            =   2400
            TabIndex        =   67
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Código del Cliente"
            Height          =   240
            Left            =   120
            TabIndex        =   66
            Top             =   360
            Width           =   1575
         End
         Begin MSForms.TextBox txt_contrato 
            DataField       =   "contrato_num"
            DataSource      =   "clientes"
            Height          =   300
            Left            =   2400
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   600
            Width           =   2655
            VariousPropertyBits=   746604575
            Size            =   "4683;529"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.TextBox txt_clientes 
            Bindings        =   "frm_pedidos.frx":08CA
            DataField       =   "codigo"
            DataSource      =   "clientes"
            Height          =   300
            Left            =   120
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   600
            Width           =   2175
            VariousPropertyBits=   746604575
            Size            =   "3836;529"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label Label3 
            Caption         =   "Apellidos y Nombres"
            Height          =   240
            Left            =   120
            TabIndex        =   63
            Top             =   960
            Width           =   1575
         End
         Begin MSForms.TextBox txt_nombre 
            Bindings        =   "frm_pedidos.frx":0909
            DataField       =   "cliente"
            DataSource      =   "clientes"
            Height          =   300
            Left            =   120
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   1200
            Width           =   4935
            VariousPropertyBits=   746604571
            Size            =   "8705;529"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label Label4 
            Caption         =   "Dirección"
            Height          =   240
            Left            =   120
            TabIndex        =   61
            Top             =   2160
            Width           =   1575
         End
         Begin MSForms.TextBox txt_direccion 
            Bindings        =   "frm_pedidos.frx":095B
            DataField       =   "direccion"
            DataSource      =   "clientes"
            Height          =   540
            Left            =   120
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   2400
            Width           =   4935
            VariousPropertyBits=   -1400879073
            Size            =   "8705;952"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label Label5 
            Caption         =   "Cédula de Identidad"
            Height          =   240
            Left            =   120
            TabIndex        =   59
            Top             =   1560
            Width           =   1575
         End
         Begin MSForms.TextBox TextBox4 
            Bindings        =   "frm_pedidos.frx":09A6
            DataField       =   "cedula"
            DataSource      =   "clientes"
            Height          =   300
            Left            =   120
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   1800
            Width           =   2175
            VariousPropertyBits=   746604571
            Size            =   "3836;529"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label Label6 
            Caption         =   "Zona"
            Height          =   240
            Left            =   2400
            TabIndex        =   57
            Top             =   1560
            Width           =   615
         End
         Begin MSForms.TextBox txt_zona 
            Bindings        =   "frm_pedidos.frx":09D4
            DataField       =   "id_ruta"
            DataSource      =   "clientes"
            Height          =   300
            Left            =   2400
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   1800
            Width           =   615
            VariousPropertyBits=   746604575
            Size            =   "1085;529"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label Label10 
            Caption         =   "Teléfono Habitación"
            Height          =   240
            Left            =   120
            TabIndex        =   55
            Top             =   3000
            Width           =   1575
         End
         Begin MSForms.TextBox TextBox10 
            Bindings        =   "frm_pedidos.frx":09FC
            DataField       =   "telefono_hab"
            DataSource      =   "clientes"
            Height          =   300
            Left            =   120
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   3240
            Width           =   1695
            VariousPropertyBits=   746604575
            Size            =   "2990;529"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label Label11 
            Caption         =   "Teléfono Celular"
            Height          =   240
            Left            =   2040
            TabIndex        =   53
            Top             =   3000
            Width           =   1455
         End
         Begin MSForms.TextBox TextBox11 
            Bindings        =   "frm_pedidos.frx":0A24
            DataField       =   "telefono_cel"
            DataSource      =   "clientes"
            Height          =   300
            Left            =   2040
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   3240
            Width           =   2175
            VariousPropertyBits=   746604575
            Size            =   "3836;529"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
      End
      Begin VB.Frame Frame33 
         Caption         =   "PEDIDOS"
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
         Height          =   3735
         Left            =   5280
         TabIndex        =   44
         Top             =   120
         Width           =   3255
         Begin MSDataListLib.DataCombo dcmb_kgs 
            Bindings        =   "frm_pedidos.frx":0A4C
            Height          =   315
            Left            =   240
            TabIndex        =   1
            Top             =   1200
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "id_inst"
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
         Begin VB.CommandButton cmd_procesar 
            Caption         =   "Procesar"
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
            Height          =   615
            Left            =   240
            TabIndex        =   4
            Top             =   2880
            Width           =   2895
         End
         Begin MSForms.TextBox txt_precio_cilind 
            Bindings        =   "frm_pedidos.frx":0A66
            CausesValidation=   0   'False
            DataField       =   "precio_cilindro"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "##,##"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "instalacion"
            Height          =   300
            Left            =   1320
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1815
            VariousPropertyBits=   746604575
            BackColor       =   -2147483633
            Size            =   "3201;529"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label8 
            Caption         =   "Kgs"
            Height          =   240
            Left            =   240
            TabIndex        =   78
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "Precio de Venta"
            Height          =   240
            Left            =   1320
            TabIndex        =   77
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label Label12 
            Caption         =   "Cantidad"
            Height          =   225
            Left            =   240
            TabIndex        =   50
            Top             =   1560
            Width           =   855
         End
         Begin MSForms.TextBox txt_cantidad 
            Bindings        =   "frm_pedidos.frx":0A76
            DataField       =   "telefono_hab"
            DataSource      =   "Ado_Clientes"
            Height          =   300
            Left            =   240
            TabIndex        =   2
            Top             =   1800
            Width           =   975
            VariousPropertyBits=   746604569
            Size            =   "1720;529"
            Value           =   "1"
            FontEffects     =   1073750017
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin VB.Label Label13 
            Caption         =   "Precio del Cilindro"
            Height          =   240
            Left            =   1320
            TabIndex        =   49
            Top             =   960
            Width           =   1335
         End
         Begin MSForms.TextBox txt_precio 
            Bindings        =   "frm_pedidos.frx":0A98
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "##,##"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            Height          =   300
            Left            =   1320
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   1800
            Width           =   1815
            VariousPropertyBits=   746604575
            BackColor       =   -2147483633
            Size            =   "3201;529"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label14 
            Caption         =   "Fecha de Entrega"
            Height          =   240
            Left            =   240
            TabIndex        =   48
            Top             =   2160
            Width           =   1575
         End
         Begin MSForms.TextBox txt_fecha_entrega 
            Bindings        =   "frm_pedidos.frx":0AB3
            CausesValidation=   0   'False
            Height          =   300
            Left            =   240
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   2400
            Width           =   2895
            VariousPropertyBits=   746604575
            BackColor       =   -2147483633
            Size            =   "5106;529"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.TextBox txt_fecha_pedido 
            Bindings        =   "frm_pedidos.frx":0ACE
            CausesValidation=   0   'False
            Height          =   300
            Left            =   240
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   600
            Width           =   2895
            VariousPropertyBits=   746604575
            BackColor       =   -2147483633
            Size            =   "5106;529"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin VB.Label Label15 
            Caption         =   "Fecha de Pedido"
            Height          =   240
            Left            =   240
            TabIndex        =   45
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "CONTROL DE DESPACHO"
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
         Height          =   6015
         Left            =   8640
         TabIndex        =   16
         Top             =   120
         Width           =   2775
         Begin VB.TextBox Txt_total 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            CausesValidation=   0   'False
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   5640
            Width           =   495
         End
         Begin VB.Frame Frame10 
            Caption         =   "ZONA Nº 4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   120
            TabIndex        =   68
            Top             =   4320
            Width           =   2535
            Begin VB.TextBox Txt_total_zona4 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               CausesValidation=   0   'False
               DataField       =   "tot_zona4"
               DataSource      =   "despachototal4"
               Height          =   285
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   95
               TabStop         =   0   'False
               Top             =   840
               Width           =   495
            End
            Begin VB.TextBox txt_c_z_4_43 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               CausesValidation=   0   'False
               DataField       =   "cant_43_z4"
               DataSource      =   "despachototal4"
               Height          =   285
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   72
               TabStop         =   0   'False
               Top             =   480
               Width           =   495
            End
            Begin VB.TextBox txt_c_z_4_27 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               CausesValidation=   0   'False
               DataField       =   "cant_27_z4"
               DataSource      =   "despachototal4"
               Height          =   285
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   71
               TabStop         =   0   'False
               Top             =   480
               Width           =   495
            End
            Begin VB.TextBox txt_c_z_4_18 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               CausesValidation=   0   'False
               DataField       =   "cant_18_z4"
               DataSource      =   "despachototal4"
               Height          =   285
               Left            =   720
               Locked          =   -1  'True
               TabIndex        =   70
               TabStop         =   0   'False
               Top             =   480
               Width           =   495
            End
            Begin VB.TextBox txt_c_z_4_10 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               CausesValidation=   0   'False
               DataField       =   "cant_10_z4"
               DataSource      =   "despachototal4"
               Height          =   285
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   69
               TabStop         =   0   'False
               Top             =   480
               Width           =   495
            End
            Begin VB.Label Lbl_sub_total4 
               Caption         =   "Sub-Total:"
               Height          =   255
               Left            =   360
               TabIndex        =   93
               Top             =   840
               Width           =   855
            End
            Begin VB.Label Label32 
               Caption         =   "43 kgs"
               Height          =   255
               Left            =   1920
               TabIndex        =   76
               Top             =   240
               Width           =   495
            End
            Begin VB.Label Label31 
               Caption         =   "27 kgs"
               Height          =   255
               Left            =   1320
               TabIndex        =   75
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label30 
               Caption         =   "18 kgs"
               Height          =   255
               Left            =   720
               TabIndex        =   74
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label29 
               Caption         =   "10 kgs"
               Height          =   255
               Left            =   120
               TabIndex        =   73
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "ZONA Nº 1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   120
            TabIndex        =   35
            Top             =   360
            Width           =   2535
            Begin VB.TextBox Txt_total_zona1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               CausesValidation=   0   'False
               DataField       =   "tot_zona1"
               DataSource      =   "despachototal1"
               Height          =   285
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   88
               TabStop         =   0   'False
               Top             =   840
               Width           =   495
            End
            Begin VB.TextBox txt_c_z_1_10 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               CausesValidation=   0   'False
               DataField       =   "cant_10_z1"
               DataSource      =   "despachototal1"
               Height          =   285
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   480
               Width           =   495
            End
            Begin VB.TextBox txt_c_z_1_18 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               CausesValidation=   0   'False
               DataField       =   "cant_18_z1"
               DataSource      =   "despachototal1"
               Height          =   285
               Left            =   720
               Locked          =   -1  'True
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   480
               Width           =   495
            End
            Begin VB.TextBox txt_c_z_1_27 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               CausesValidation=   0   'False
               DataField       =   "cant_27_z1"
               DataSource      =   "despachototal1"
               Height          =   285
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   480
               Width           =   495
            End
            Begin VB.TextBox txt_c_z_1_43 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               CausesValidation=   0   'False
               DataField       =   "cant_43_z1"
               DataSource      =   "despachototal1"
               Height          =   285
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   480
               Width           =   495
            End
            Begin VB.Label Lbl_sub_total1 
               Caption         =   "Sub-Total:"
               Height          =   255
               Left            =   360
               TabIndex        =   89
               Top             =   840
               Width           =   855
            End
            Begin VB.Label Label16 
               Caption         =   "10 kgs"
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label17 
               Caption         =   "18 kgs"
               Height          =   255
               Left            =   720
               TabIndex        =   42
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label18 
               Caption         =   "27 kgs"
               Height          =   255
               Left            =   1320
               TabIndex        =   41
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label19 
               Caption         =   "43 kgs"
               Height          =   255
               Left            =   1920
               TabIndex        =   40
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "ZONA Nº 2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   120
            TabIndex        =   26
            Top             =   1680
            Width           =   2535
            Begin VB.TextBox Txt_total_zona2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               CausesValidation=   0   'False
               DataField       =   "tot_zona2"
               DataSource      =   "despachototal2"
               Height          =   285
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   90
               TabStop         =   0   'False
               Top             =   840
               Width           =   495
            End
            Begin VB.TextBox txt_c_z_2_43 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               CausesValidation=   0   'False
               DataField       =   "cant_43_z2"
               DataSource      =   "despachototal2"
               Height          =   285
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   480
               Width           =   495
            End
            Begin VB.TextBox txt_c_z_2_27 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               CausesValidation=   0   'False
               DataField       =   "cant_27_z2"
               DataSource      =   "despachototal2"
               Height          =   285
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   480
               Width           =   495
            End
            Begin VB.TextBox txt_c_z_2_18 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               CausesValidation=   0   'False
               DataField       =   "cant_18_z2"
               DataSource      =   "despachototal2"
               Height          =   285
               Left            =   720
               Locked          =   -1  'True
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   480
               Width           =   495
            End
            Begin VB.TextBox txt_c_z_2_10 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               CausesValidation=   0   'False
               DataField       =   "cant_10_z2"
               DataSource      =   "despachototal2"
               Height          =   285
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   480
               Width           =   495
            End
            Begin VB.Label Lbl_sub_total2 
               Caption         =   "Sub-Total:"
               Height          =   255
               Left            =   360
               TabIndex        =   91
               Top             =   840
               Width           =   855
            End
            Begin VB.Label Label20 
               Caption         =   "43 kgs"
               Height          =   255
               Left            =   1920
               TabIndex        =   34
               Top             =   240
               Width           =   495
            End
            Begin VB.Label Label21 
               Caption         =   "27 kgs"
               Height          =   255
               Left            =   1320
               TabIndex        =   33
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label22 
               Caption         =   "18 kgs"
               Height          =   255
               Left            =   720
               TabIndex        =   32
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label23 
               Caption         =   "10 kgs"
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "ZONA Nº 3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   120
            TabIndex        =   17
            Top             =   3000
            Width           =   2535
            Begin VB.TextBox Txt_total_zona3 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               CausesValidation=   0   'False
               DataField       =   "tot_zona3"
               DataSource      =   "despachototal3"
               Height          =   285
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   94
               TabStop         =   0   'False
               Top             =   840
               Width           =   495
            End
            Begin VB.TextBox txt_c_z_3_10 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               CausesValidation=   0   'False
               DataField       =   "cant_10_z3"
               DataSource      =   "despachototal3"
               Height          =   285
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   480
               Width           =   495
            End
            Begin VB.TextBox txt_c_z_3_18 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               CausesValidation=   0   'False
               DataField       =   "cant_18_z3"
               DataSource      =   "despachototal3"
               Height          =   285
               Left            =   720
               Locked          =   -1  'True
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   480
               Width           =   495
            End
            Begin VB.TextBox txt_c_z_3_27 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               CausesValidation=   0   'False
               DataField       =   "cant_27_z3"
               DataSource      =   "despachototal3"
               Height          =   285
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   480
               Width           =   495
            End
            Begin VB.TextBox txt_c_z_3_43 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               CausesValidation=   0   'False
               DataField       =   "cant_43_z3"
               DataSource      =   "despachototal3"
               Height          =   285
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   480
               Width           =   495
            End
            Begin VB.Label Lbl_sub_total3 
               Caption         =   "Sub-Total:"
               Height          =   255
               Left            =   360
               TabIndex        =   92
               Top             =   840
               Width           =   855
            End
            Begin VB.Label Label24 
               Caption         =   "10 kgs"
               Height          =   255
               Left            =   120
               TabIndex        =   25
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label25 
               Caption         =   "18 kgs"
               Height          =   255
               Left            =   720
               TabIndex        =   24
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label26 
               Caption         =   "27 kgs"
               Height          =   255
               Left            =   1320
               TabIndex        =   23
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label27 
               Caption         =   "43 kgs"
               Height          =   255
               Left            =   1920
               TabIndex        =   22
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Label Lbl_total 
            Caption         =   "Total:"
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
            Left            =   720
            TabIndex        =   96
            Top             =   5640
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "HISTÓRICO DE PEDIDOS DEL CLIENTE"
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
         Height          =   2175
         Left            =   0
         TabIndex        =   14
         Top             =   3960
         Width           =   8535
         Begin MSDataGridLib.DataGrid DGrid_pedidos 
            Bindings        =   "frm_pedidos.frx":0AE9
            Height          =   1815
            Left            =   120
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   240
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   3201
            _Version        =   393216
            AllowUpdate     =   0   'False
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
            ColumnCount     =   7
            BeginProperty Column00 
               DataField       =   "id_pedido"
               Caption         =   "id_pedido"
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
               DataField       =   "status"
               Caption         =   "status"
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
               DataField       =   "fecha_pedido"
               Caption         =   "fecha_pedido"
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
               DataField       =   "fecha_cancel"
               Caption         =   "fecha_cancel"
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
               DataField       =   "monto_fac"
               Caption         =   "monto_fac"
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
               DataField       =   "cant_pedido"
               Caption         =   "cant_pedido"
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
               DataField       =   "usuario_liq"
               Caption         =   "usuario_liq"
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
               BeginProperty Column00 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   540,284
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1110,047
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1184,882
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1365,165
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   989,858
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   1739,906
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3960
      TabIndex        =   11
      Top             =   1080
      Width           =   7455
      Begin VB.CommandButton Busquedad_avanzadas 
         Caption         =   "Búsqueda Avanzada"
         Height          =   375
         Index           =   0
         Left            =   5400
         TabIndex        =   12
         Tag             =   "Lista todos los inmuebles registrados"
         Top             =   0
         Width           =   1935
      End
      Begin MSDataListLib.DataCombo Dcmb_Buscar 
         Bindings        =   "frm_pedidos.frx":0B15
         Height          =   315
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Pulse doble click para cambiar el tipo de busqueda, después de presionar búsqueda avanzada"
         Top             =   0
         Width           =   5175
         _ExtentX        =   9128
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
   End
   Begin MSAdodcLib.Adodc instalacion 
      Height          =   375
      Left            =   12840
      Top             =   5280
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
   Begin MSAdodcLib.Adodc pedidos 
      Height          =   375
      Left            =   12840
      Top             =   5640
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
   Begin MSAdodcLib.Adodc hist_pedidos 
      Height          =   375
      Left            =   2040
      Top             =   360
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
      CommandType     =   1
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
      RecordSource    =   "select * from tbl_pedidos where  codigo =''"
      Caption         =   "hist_pedidos"
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
   Begin MSAdodcLib.Adodc despacho 
      Height          =   375
      Left            =   12840
      Top             =   6000
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
      RecordSource    =   "tbl_despacho"
      Caption         =   "despacho"
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
   Begin MSAdodcLib.Adodc chofer 
      Height          =   375
      Left            =   12840
      Top             =   6360
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
      RecordSource    =   "tbl_chofer"
      Caption         =   "chofer"
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
   Begin MSAdodcLib.Adodc despachototal1 
      Height          =   375
      Left            =   12840
      Top             =   8040
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
      CommandType     =   8
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
      RecordSource    =   "select * from tbl_despacho where fecha_pedido=''"
      Caption         =   "despachototal1"
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
   Begin MSAdodcLib.Adodc despachototal2 
      Height          =   375
      Left            =   12840
      Top             =   7680
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
      CommandType     =   8
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
      RecordSource    =   "select * from tbl_despacho where fecha_pedido=''"
      Caption         =   "despachototal2"
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
   Begin MSAdodcLib.Adodc despachototal3 
      Height          =   375
      Left            =   12840
      Top             =   4920
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
      CommandType     =   8
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
      RecordSource    =   "select * from tbl_despacho where fecha_pedido=''"
      Caption         =   "despachototal3"
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
   Begin MSAdodcLib.Adodc despachototal4 
      Height          =   375
      Left            =   12840
      Top             =   4560
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
      CommandType     =   8
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
      RecordSource    =   "select * from tbl_despacho where fecha_pedido=''"
      Caption         =   "despachototal4"
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
   Begin MSAdodcLib.Adodc despachototal 
      Height          =   375
      Left            =   12840
      Top             =   8760
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
      CommandType     =   8
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
      RecordSource    =   "select * from tbl_despacho where fecha_pedido=''"
      Caption         =   "despachototal"
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
   Begin MSAdodcLib.Adodc liquidado 
      Height          =   375
      Left            =   12840
      Top             =   6720
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
      RecordSource    =   "tbl_liquidado"
      Caption         =   "liquidado"
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
      Left            =   2040
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
   Begin MSAdodcLib.Adodc resumen_mensual_ventas 
      Height          =   375
      Left            =   12840
      Top             =   8400
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
      RecordSource    =   "tbl_resumen_mensual_ventas"
      Caption         =   "resumen_mensual_ventas"
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
   Begin MSAdodcLib.Adodc res_inventario 
      Height          =   375
      Left            =   12840
      Top             =   7080
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
      RecordSource    =   "tbl_resumen_inventario"
      Caption         =   "res_inventatio"
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
   Begin MSAdodcLib.Adodc control 
      Height          =   375
      Left            =   4200
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
      RecordSource    =   "tbl_control_procesos"
      Caption         =   "control"
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
      Left            =   2880
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
      Left            =   480
      TabIndex        =   10
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
   Begin VB.Label Label_titulo 
      BackColor       =   &H80000001&
      Caption         =   "  PEDIDOS JERGAS"
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
      TabIndex        =   9
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
Attribute VB_Name = "frm_pedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim total_iva
Dim total_factura

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
        
        Call habilitar_botones(False)
        
    Else
    
        If Me.clientes.Recordset.RecordCount > 1 Then
        
            MsgBox Me.clientes.Recordset.RecordCount & " encontrados"
            
            Busq_Avanzada = True
            
        End If
        
        Call habilitar_botones(True)
        
    End If
    
Else
    
    clientes.Recordset.MoveFirst
    
    strquery = "codigo = '" & Dcmb_Buscar.Text & "'"

    clientes.Recordset.Find strquery
    
    If clientes.Recordset.EOF Then
    
            MsgBox "Código suministrado no encontrado, por favor verifique ", vbInformation, "SIAGEP"
            
            Dcmb_Buscar.Text = ""
            
            Dcmb_Buscar.SetFocus
            
            Call habilitar_botones(False)
                    
    Else
    
            Call habilitar_botones(True)
        
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
Private Sub habilitar_botones(valor As Boolean)
Me.cmd_clientes.Enabled = valor
Me.cmd_estado.Enabled = valor
'Me.cmd_ventas_materiales.Enabled = valor
'Me.cmd_liquidacion.Enabled = valor

End Sub
Private Sub buscar_asignar_zona()

Dim zona1_10, c_z_1_10, ruta, iden_chofer, strquery
'------------------------------------------------
'Este procedimiento se encarga de asignar la zona
'de un pedido en especifico
'------------------------------------------------

On Error GoTo ControlError
'--------------------------------------------------------------
'Hay que buscar el id_chofer para asignar el id_ruta (despacho)
'--------------------------------------------------------------
If (Me.txt_zona.Text <> "") Then

    Me.chofer.Recordset.MoveFirst
    
    strquery = "id_ruta = '" & txt_zona.Text & "'"

    chofer.Recordset.Find strquery
    
    If chofer.Recordset.EOF Then
    
            MsgBox "No existe Chofer para la Zona suministrada por el cliente, por favor verifique ", vbInformation, "JerGas C.A."
            Exit Sub
            
    Else
    
        iden_chofer = chofer.Recordset!id_chofer
        
    End If
    
Else

    MsgBox "No se puede asignar el despacho debido a que no hay una zona", vbCritical
    Exit Sub
    
End If
'---------------------------------------------------
'Si todo salio bien aquí tengo el id_chofer (chofer)
'con respecto a la ruta del cliente dado.
'Debemos buscar si existe registro para el dia de
'hoy, si no es así asignarlo.
'---------------------------------------------------

'despacho esta perdiendo el bookmark
'mvBookMark = despacho.Recordset.Bookmark
'
'despacho.Recordset.MoveFirst
  
despacho.CommandType = adCmdText

'El numero 1 es el chofer número uno, esto se realizará para cada chofer

'strquery = "SELECT * FROM tbl_despacho WHERE tbl_despacho.id_chofer = '" & iden_chofer & "' AND tbl_despacho.fecha_pedido= " & Format(Date, "dd/mm/yyyy") & ""

strquery = "SELECT * FROM tbl_despacho WHERE tbl_despacho.id_chofer = '" & iden_chofer & "' AND tbl_despacho.fecha_pedido= '" & Format(Date, "yyyy/mm/dd") & "'"

despacho.RecordSource = strquery

despacho.Refresh
'Es true si se va ha despachar en dicha zona por primera
'vez un día X
If despacho.Recordset.EOF Then

With despacho.Recordset

    If Not (.BOF And .EOF) Then
    
      mvBookMark = .Bookmark
            
      .MoveLast
        
    End If
    
    .AddNew
    
    !id_chofer = iden_chofer
    
    !fecha_pedido = Date
    
    !id_ruta = Me.txt_zona.Text
    
    'Estamos ubicados en la condición de un nuevo despacho
    'dependiendo que bombona se está vendiendo se tiene que
    'asignar el numero uno para dicha venta
    
    'Dependiendo de la Zona se asigna 1
        Select Case Me.txt_zona.Text

            Case "1"
            
                If dcmb_kgs.Text = 10 Then
                    !cant_10_z1 = CInt(txt_cantidad.Text)
                End If

                If dcmb_kgs.Text = 18 Then
                    !cant_18_z1 = CInt(txt_cantidad.Text)
                End If

                If dcmb_kgs.Text = 27 Then
                    !cant_27_z1 = CInt(txt_cantidad.Text)
                End If

                If dcmb_kgs.Text = 43 Then
                    !cant_43_z1 = CInt(txt_cantidad.Text)
                End If
                
                'Vamos a calcular los totales
                If IsNull(!tot_zona1) Then
                    !tot_zona1 = CInt(txt_cantidad.Text)
                Else
                    !tot_zona1 = !tot_zona1 + CInt(txt_cantidad.Text)
                End If
                
            Case "2"
            
                If dcmb_kgs.Text = 10 Then
                    !cant_10_z2 = CInt(txt_cantidad.Text)
                End If

                If dcmb_kgs.Text = 18 Then
                    !cant_18_z2 = CInt(txt_cantidad.Text)
                End If

                If dcmb_kgs.Text = 27 Then
                    !cant_27_z2 = CInt(txt_cantidad.Text)
                End If

                If dcmb_kgs.Text = 43 Then
                    !cant_43_z2 = CInt(txt_cantidad.Text)
                End If
                
                'Vamos a calcular los totales
                If IsNull(!tot_zona2) Then
                    !tot_zona2 = CInt(txt_cantidad.Text)
                Else
                    !tot_zona2 = !tot_zona2 + CInt(txt_cantidad.Text)
                End If
                
            Case "3"
            
                If dcmb_kgs.Text = 10 Then
                    !cant_10_z3 = CInt(txt_cantidad.Text)
                End If

                If dcmb_kgs.Text = 18 Then
                    !cant_18_z3 = CInt(txt_cantidad.Text)
                End If

                If dcmb_kgs.Text = 27 Then
                    !cant_27_z3 = CInt(txt_cantidad.Text)
                End If

                If dcmb_kgs.Text = 43 Then
                    !cant_43_z3 = CInt(txt_cantidad.Text)
                End If
                
                'Vamos a calcular los totales
                If IsNull(!tot_zona3) Then
                    !tot_zona3 = CInt(txt_cantidad.Text)
                Else
                    !tot_zona3 = !tot_zona3 + CInt(txt_cantidad.Text)
                End If
                
            Case "4"
            
                If dcmb_kgs.Text = 10 Then
                    !cant_10_z4 = CInt(txt_cantidad.Text)
                End If

                If dcmb_kgs.Text = 18 Then
                    !cant_18_z4 = CInt(txt_cantidad.Text)
                End If

                If dcmb_kgs.Text = 27 Then
                    !cant_27_z4 = CInt(txt_cantidad.Text)
                End If

                If dcmb_kgs.Text = 43 Then
                    !cant_43_z4 = CInt(txt_cantidad.Text)
                End If
                
                'Vamos a calcular los totales
                If IsNull(!tot_zona4) Then
                    !tot_zona4 = CInt(txt_cantidad.Text)
                Else
                    !tot_zona4 = !tot_zona4 + CInt(txt_cantidad.Text)
                End If
                
        End Select
                
    .Update

End With


Else

    '---------------------------------------------
    'Si id_chofer ya existe para dicha id_ruta
    'entonces, se debe ir modificando las cantidad
    'de pedidos para dicho chofer, y totalizando el
    'numero de bombonas, no deberia ser mas de 60
    'por camión
    '---------------------------------------------
    With despacho.Recordset

        
'        !id_chofer = iden_chofer
'
'        !fecha_pedido = Date
'
'        !id_ruta = Me.txt_zona.Text
        
        'Dependiendo de la Zona se suma y se asigna
        
        Select Case Me.txt_zona.Text
        
            Case "1"
                If dcmb_kgs.Text = 10 Then
                    If IsNull(!cant_10_z1) Then
                        !cant_10_z1 = CInt(txt_cantidad.Text)
                    Else
                        !cant_10_z1 = !cant_10_z1 + CInt(txt_cantidad.Text)
                    End If
                End If
                If dcmb_kgs.Text = 18 Then
                    If IsNull(!cant_18_z1) Then
                        !cant_18_z1 = CInt(txt_cantidad.Text)
                    Else
                        !cant_18_z1 = !cant_18_z1 + CInt(txt_cantidad.Text)
                    End If
                End If
                If dcmb_kgs.Text = 27 Then
                    If IsNull(!cant_27_z1) Then
                        !cant_27_z1 = CInt(txt_cantidad.Text)
                    Else
                        !cant_27_z1 = !cant_27_z1 + CInt(txt_cantidad.Text)
                    End If
                End If
                If dcmb_kgs.Text = 43 Then
                    If IsNull(!cant_43_z1) Then
                        !cant_43_z1 = CInt(txt_cantidad.Text)
                    Else
                        !cant_43_z1 = !cant_43_z1 + CInt(txt_cantidad.Text)
                    End If
                End If
                
                'Vamos a calcular los totales
                If IsNull(!tot_zona1) Then
                        !tot_zona1 = CInt(txt_cantidad.Text)
                Else
                        !tot_zona1 = !tot_zona1 + CInt(txt_cantidad.Text)
                End If
                
                
            Case "2"
                If dcmb_kgs.Text = 10 Then
                    If IsNull(!cant_10_z2) Then
                        !cant_10_z2 = CInt(txt_cantidad.Text)
                    Else
                        !cant_10_z2 = !cant_10_z2 + CInt(txt_cantidad.Text)
                    End If
                End If
                If dcmb_kgs.Text = 18 Then
                    If IsNull(!cant_18_z2) Then
                        !cant_18_z2 = CInt(txt_cantidad.Text)
                    Else
                        !cant_18_z2 = !cant_18_z2 + CInt(txt_cantidad.Text)
                    End If
                End If
                If dcmb_kgs.Text = 27 Then
                    If IsNull(!cant_27_z2) Then
                        !cant_27_z2 = CInt(txt_cantidad.Text)
                    Else
                        !cant_27_z2 = !cant_27_z2 + CInt(txt_cantidad.Text)
                    End If
                End If
                If dcmb_kgs.Text = 43 Then
                    If IsNull(!cant_43_z2) Then
                        !cant_43_z2 = CInt(txt_cantidad.Text)
                    Else
                        !cant_43_z2 = !cant_43_z2 + CInt(txt_cantidad.Text)
                    End If
                End If
                'Vamos a calcular los totales
                If IsNull(!tot_zona2) Then
                        !tot_zona2 = CInt(txt_cantidad.Text)
                Else
                        !tot_zona2 = !tot_zona2 + CInt(txt_cantidad.Text)
                End If
                
            Case "3"
                If dcmb_kgs.Text = 10 Then
                    If IsNull(!cant_10_z3) Then
                        !cant_10_z3 = CInt(txt_cantidad.Text)
                    Else
                        !cant_10_z3 = !cant_10_z3 + CInt(txt_cantidad.Text)
                    End If
                End If
                If dcmb_kgs.Text = 18 Then
                    If IsNull(!cant_18_z3) Then
                        !cant_18_z3 = CInt(txt_cantidad.Text)
                    Else
                        !cant_18_z3 = !cant_18_z3 + CInt(txt_cantidad.Text)
                    End If
                End If
                If dcmb_kgs.Text = 27 Then
                    If IsNull(!cant_27_z3) Then
                        !cant_27_z3 = CInt(txt_cantidad.Text)
                    Else
                        !cant_27_z3 = !cant_27_z3 + CInt(txt_cantidad.Text)
                    End If
                End If
                If dcmb_kgs.Text = 43 Then
                    If IsNull(!cant_43_z3) Then
                        !cant_43_z3 = CInt(txt_cantidad.Text)
                    Else
                        !cant_43_z3 = !cant_43_z3 + CInt(txt_cantidad.Text)
                    End If
                End If
                'Vamos a calcular los totales
                If IsNull(!tot_zona3) Then
                        !tot_zona3 = CInt(txt_cantidad.Text)
                Else
                        !tot_zona3 = !tot_zona3 + CInt(txt_cantidad.Text)
                End If

                
            Case "4"
                If dcmb_kgs.Text = 10 Then
                    If IsNull(!cant_10_z4) Then
                        !cant_10_z4 = CInt(txt_cantidad.Text)
                    Else
                        !cant_10_z4 = !cant_10_z4 + CInt(txt_cantidad.Text)
                    End If
                End If
                If dcmb_kgs.Text = 18 Then
                    If IsNull(!cant_18_z4) Then
                        !cant_18_z4 = CInt(txt_cantidad.Text)
                    Else
                        !cant_18_z4 = !cant_18_z4 + CInt(txt_cantidad.Text)
                    End If
                End If
                If dcmb_kgs.Text = 27 Then
                    If IsNull(!cant_27_z4) Then
                        !cant_27_z4 = CInt(txt_cantidad.Text)
                    Else
                        !cant_27_z4 = !cant_27_z4 + CInt(txt_cantidad.Text)
                    End If
                End If
                If dcmb_kgs.Text = 43 Then
                    If IsNull(!cant_43_z4) Then
                        !cant_43_z4 = CInt(txt_cantidad.Text)
                    Else
                        !cant_43_z4 = !cant_43_z4 + CInt(txt_cantidad.Text)
                    End If
                End If
                'Vamos a calcular los totales
                If IsNull(!tot_zona4) Then
                        !tot_zona4 = CInt(txt_cantidad.Text)
                Else
                        !tot_zona4 = !tot_zona4 + CInt(txt_cantidad.Text)
                End If
                'Vamos a calcular los totales

        End Select
''        'Vamos a calcular el total General
''        If IsNull(!total_despacho) Then
''                !total_despacho = 1
''        Else
''                !total_despacho = !total_despacho + 1
''        End If
        
        .Update
        
    End With
    
End If

'Cierro la conexión
despacho.Recordset.Close

'---------------------------------------------
'Verifica el numero de pedidos por chofer
'no debe ser mayor de 60, para la fecha actual
'---------------------------------------------

despacho.CommandType = adCmdText
'Se tiene que totalizar los pedidos por chofer
'El numero 1 es el chofer número uno, esto se realizará para cada chofer

despacho.RecordSource = "SELECT COUNT(cant_10_z1) AS c_z_1_10  FROM tbl_despacho WHERE tbl_despacho.id_chofer = " & iden_chofer & " AND tbl_despacho.fecha_pedido= " & Date & "  "

despacho.Refresh

If despacho.Recordset.EOF Then

    c_z_1_10 = despacho.Recordset!c_z_1_10

End If

despacho.Recordset.Close
'------------------------------------------------------------

'despacho.CommandType = adCmdText
'
'despacho.RecordSource = "SELECT * FROM tbl_despacho WHERE tbl_despacho.codigo = '" & Me.txt_clientes.Text & "'"
'
'despacho.Refresh

Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "JerGas")
        
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
Me.cmd_eliminar.Enabled = False

Screen.MousePointer = 11

If DGrid_pedidos.SelBookmarks.Count = 0 Then
    
    MsgBox "No se hallaron Pedidos marcados para Eliminar."
    Me.cmd_eliminar.Enabled = True
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
       FGNRO_LIQ_RESTA_PEDIDO
       FGNRO_LIQ_RESTA_FACTURA
       FGNRO_LIQ_RESTA_CONTROL
       End If
    
hist_pedidos.Refresh
pedidos.Refresh
Screen.MousePointer = 0

Exit Sub

control_error:
Screen.MousePointer = 0
    MsgBox Err.Description

End Sub

Private Sub cmd_ventas_materiales_Click()
frm_ventas_materiales.Show
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
        
        Call habilitar_botones(False)
        
    Else
    
        If Me.clientes.Recordset.RecordCount > 1 Then
        
            MsgBox Me.clientes.Recordset.RecordCount & " encontrados"
            
            Busq_Avanzada = True
            
        End If
        
        Call habilitar_botones(True)
        
    End If
    
Else
    
    clientes.Recordset.MoveFirst
    
    strquery = "contrato_num = '" & Dcmb_Buscar.Text & "'"

    clientes.Recordset.Find strquery
    
    If clientes.Recordset.EOF Then
    
            MsgBox "Nº de Contrato suministrado no encontrado, por favor verifique ", vbInformation, "SIAGEP"
            
            Dcmb_Buscar.Text = ""
            
            Dcmb_Buscar.SetFocus
            
            Call habilitar_botones(False)
                    
    Else
    
            Call habilitar_botones(True)
        
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
        
        Call habilitar_botones(False)
        
    Else
    
        If Me.clientes.Recordset.RecordCount > 1 Then
        
            MsgBox Me.clientes.Recordset.RecordCount & " encontrados"
            
            Busq_Avanzada = True
            
        End If
        
        Call habilitar_botones(True)
        
    End If
    
Else
    
    clientes.Recordset.MoveFirst
    
    strquery = "cedula = '" & Dcmb_Buscar.Text & "'"

    clientes.Recordset.Find strquery
    
    If clientes.Recordset.EOF Then
    
            MsgBox "Cédula del Cliente suministrado no encontrado, por favor verifique ", vbInformation, "SIAGEP"
            
            Dcmb_Buscar.Text = ""
            
            Dcmb_Buscar.SetFocus
            
            Call habilitar_botones(False)
                    
    Else
    
            Call habilitar_botones(True)
        
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
        
        Call habilitar_botones(False)
        
    Else
    
        If Me.clientes.Recordset.RecordCount > 1 Then
        
            MsgBox Me.clientes.Recordset.RecordCount & " encontrados"
            
            Busq_Avanzada = True
            
        End If
        
        Call habilitar_botones(True)
        
    End If
    
Else
    
    clientes.Recordset.MoveFirst
    
    strquery = "cliente = '" & Dcmb_Buscar.Text & "'"

    clientes.Recordset.Find strquery
    
    If clientes.Recordset.EOF Then
    
            MsgBox "Nombre del Cliente suministrado no encontrado, por favor verifique ", vbInformation, "JerGas"
            
            Dcmb_Buscar.Text = ""
            
            Dcmb_Buscar.SetFocus
            
            Call habilitar_botones(False)
                    
    Else
    
            Call habilitar_botones(True)
        
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
        
        Call habilitar_botones(False)
        
    Else
    
        If Me.clientes.Recordset.RecordCount > 1 Then
        
            MsgBox Me.clientes.Recordset.RecordCount & " encontrados"
            
            Busq_Avanzada = True
            
        End If
        
        Call habilitar_botones(True)
        
    End If
    
Else
    
    clientes.Recordset.MoveFirst
    
    strquery = "direccion = '" & Dcmb_Buscar.Text & "'"

    clientes.Recordset.Find strquery
    
    If clientes.Recordset.EOF Then
    
            MsgBox "Dirección del Cliente suministrado no encontrado, por favor verifique ", vbInformation, "JerGas"
            
            Dcmb_Buscar.Text = ""
            
            Dcmb_Buscar.SetFocus
            
            Call habilitar_botones(False)
                    
    Else
    
            Call habilitar_botones(True)
        
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
'End If
End Sub

Private Sub dcmb_kgs_Click(Area As Integer)
On Error GoTo ControlError
If (Area = 2) Then
    instalacion.Recordset.MoveFirst
    
    strquery = "id_inst = '" & dcmb_kgs.Text & "'"

    instalacion.Recordset.Find strquery
    
    If instalacion.Recordset.EOF Then
    
            dcmb_kgs.Text = ""
    Else
        
        Me.txt_cantidad.Enabled = True
        
    End If
End If
Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "JerGas")
        End Select
End Sub

Private Sub DGrid_historico_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_clientes.FontBold = False
Me.cmd_estado.FontBold = False
Me.cmd_liquidacion.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_ventas_materiales.FontBold = False
End Sub

Private Sub DGrid_pedidos_Click()
    Me.cmd_liquidacion.Enabled = True
    Me.cmd_eliminar.Enabled = True
    For Each Var In DGrid_pedidos.SelBookmarks
    DGrid_pedidos.Bookmark = Var
    'Si status es CA
    '---------------
    If DGrid_pedidos.Columns(1) = "CA" Then
            MsgBox "Factura ya está cancelada", vbInformation, "JerGas"
            DGrid_pedidos.SelBookmarks.Remove (DGrid_pedidos.SelBookmarks.Count - 1)
'            If DGrid_inm_liq.SelBookmarks.Count = 0 Then
'                Tex_Cuotas.Text = ""
'                Tex_Monto.Text = ""
'            End If
            Exit For
            
    End If
    Next
End Sub

Private Sub Form_Load()
Call actualizar_cn("SQL Server")
Me.txt_fecha_pedido = Date
Me.txt_fecha_entrega = DateAdd("d", 1, Date)
Me.txt_date = Date
End Sub

Private Sub Form_Resize()
Call Mover_der(Me, Label_titulo, 0)
Call Mover_centrado(Me, Frame1)
Call Mover_der(Me, Frame3, 10)
Call Mover_der(Me, Me.LABEL_BUSCA, Frame3.Width + 15)
'Call Mover_der(Me, Me.Busquedad_avanzadas, 3)
'Call Mover_der(Me, Me.Dcmb_Buscarbif, 10)
'Call Mover_der(Me, Me.LABEL_BUSCA, Me.Dcmb_Buscarbif.Width + 15)
Shape1.Width = Me.Width
Shape1.Left = 0
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_clientes.FontBold = False
Me.cmd_estado.FontBold = False
Me.cmd_liquidacion.FontBold = False
Me.cmd_procesar.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_ventas_materiales.FontBold = False
End Sub

Private Sub Frame10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_clientes.FontBold = False
Me.cmd_estado.FontBold = False
Me.cmd_liquidacion.FontBold = False
Me.cmd_procesar.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_ventas_materiales.FontBold = False
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_clientes.FontBold = False
Me.cmd_estado.FontBold = False
Me.cmd_liquidacion.FontBold = False
Me.cmd_procesar.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_ventas_materiales.FontBold = False
End Sub

Private Sub Frame33_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_procesar.FontBold = False
End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_clientes.FontBold = False
Me.cmd_estado.FontBold = False
Me.cmd_liquidacion.FontBold = False
Me.cmd_procesar.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_ventas_materiales.FontBold = False
End Sub

Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_clientes.FontBold = False
Me.cmd_estado.FontBold = False
Me.cmd_liquidacion.FontBold = False
Me.cmd_procesar.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_ventas_materiales.FontBold = False
End Sub

Private Sub TextBox14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_procesar.FontBold = False
End Sub

Private Sub Label35_Click()

End Sub

Private Sub Timer1_Timer()
On Error GoTo ControlError
Dim ttzone, ttzone1, ttzone2, ttzone3, ttzone4
    'Zona 1
    Me.despachototal1.CommandType = adCmdText
    
    Me.despachototal1.RecordSource = "SELECT * FROM tbl_despacho WHERE tbl_despacho.id_chofer = '1' and tbl_despacho.fecha_pedido >= '" & Format(Date, "yyyy/mm/dd") & "'"
    
    Me.despachototal1.Refresh
    
    Me.despachototal1.Recordset.Close
    'Zona2
    Me.despachototal2.CommandType = adCmdText

    Me.despachototal2.RecordSource = "SELECT * FROM tbl_despacho WHERE tbl_despacho.id_chofer = '2' and tbl_despacho.fecha_pedido >= '" & Format(Date, "yyyy/mm/dd") & "'"

    Me.despachototal2.Refresh

    Me.despachototal2.Recordset.Close

    'Zona3
    Me.despachototal3.CommandType = adCmdText

    Me.despachototal3.RecordSource = "SELECT * FROM tbl_despacho WHERE id_chofer = '3' and tbl_despacho.fecha_pedido >= '" & Format(Date, "yyyy/mm/dd") & "'"

    Me.despachototal3.Refresh

    Me.despachototal3.Recordset.Close

    'Zona 4
    Me.despachototal4.CommandType = adCmdText

    Me.despachototal4.RecordSource = "SELECT * FROM tbl_despacho WHERE id_chofer = '4' and tbl_despacho.fecha_pedido >= '" & Format(Date, "yyyy/mm/dd") & "'"

    Me.despachototal4.Refresh

    Me.despachototal4.Recordset.Close
    
    'Debo sumar los cuatro totales
    
    If IsNull(Txt_total_zona1) Or Txt_total_zona1 = "" Then
        ttzone1 = 0
    Else
        ttzone1 = Txt_total_zona1
    End If
    If IsNull(Txt_total_zona2) Or Txt_total_zona2 = "" Then
        ttzone2 = 0
    Else
        ttzone2 = Txt_total_zona2
    End If
    If IsNull(Txt_total_zona3) Or Txt_total_zona3 = "" Then
        ttzone3 = 0
    Else
        ttzone3 = Txt_total_zona3
    End If
    If IsNull(Txt_total_zona4) Or Txt_total_zona4 = "" Then
        ttzone4 = 0
    Else
        ttzone4 = Txt_total_zona4
    End If
    
    ttzona = CInt(ttzone1) + CInt(ttzone2) + CInt(ttzone3) + CInt(ttzone4)
    
    Txt_total = ttzona

    
Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "JerGas")
        
    End Select
End Sub

Private Sub txt_cantidad_Change()
If Me.txt_cantidad <> "" Then
Me.txt_precio = Me.txt_precio_cilind.Text * Me.txt_cantidad
Me.cmd_procesar.Enabled = True
End If
End Sub

Private Sub txt_cantidad_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_cantidad_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.txt_cantidad <> "" Then
Me.txt_precio = Me.txt_precio_cilind.Text * Me.txt_cantidad
Me.cmd_procesar.Enabled = True
End If
End Sub

Private Sub cmd_cerrar_Click()
Unload Me

End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = True
Me.cmd_clientes.FontBold = False
Me.cmd_estado.FontBold = False
Me.cmd_liquidacion.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_ventas_materiales.FontBold = False
End Sub

Private Sub cmd_clientes_Click()
Screen.MousePointer = 13
frm_clientes.Show
Screen.MousePointer = 0
End Sub

Private Sub cmd_clientes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_clientes.FontBold = True
Me.cmd_estado.FontBold = False
Me.cmd_liquidacion.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_ventas_materiales.FontBold = False
End Sub

Private Sub cmd_estado_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_clientes.FontBold = False
Me.cmd_estado.FontBold = True
Me.cmd_liquidacion.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_ventas_materiales.FontBold = False
End Sub

Private Sub cmd_eliminar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_clientes.FontBold = False
Me.cmd_estado.FontBold = False
Me.cmd_liquidacion.FontBold = False
Me.cmd_eliminar.FontBold = True
Me.cmd_ventas_materiales.FontBold = False
End Sub

Private Sub cmd_liquidacion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_clientes.FontBold = False
Me.cmd_estado.FontBold = False
Me.cmd_liquidacion.FontBold = True
Me.cmd_eliminar.FontBold = False
Me.cmd_ventas_materiales.FontBold = False
End Sub

Private Sub cmd_ventas_materiales_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_clientes.FontBold = False
Me.cmd_estado.FontBold = False
Me.cmd_liquidacion.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_ventas_materiales.FontBold = True
End Sub
Private Sub cmd_procesar_Click()
On Error GoTo ControlError
'Call actualizar_cn("SQL Server")
'Se realiza el Pedido
'Se busca el ultimo numero de pedido y se genera la proxima transaccion

Gcod_planilla = FGNRO_LIQ()
'Gcod_control = FGNRO_CONTROL()
'Gcod_factura = FGNRO_FACTURA()

With pedidos.Recordset

    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
        .MoveLast
        Me.txt_control.Text = !codigo
        Me.txt_num_control.Text = !num_control
        Me.txt_num_factura.Text = !num_factura
          
          If Me.txt_clientes.Text <> Me.txt_control.Text Then
             Gcod_control = FGNRO_CONTROL()
             Gcod_factura = FGNRO_FACTURA()
          Else
             Gcod_control = Me.txt_num_control.Text
             Gcod_factura = Me.txt_num_factura.Text
          End If
        
    End If
    'Añadimos un pedido al cliente
    .AddNew
    
    !usuario_liq = Usuario
    !id_pedido = Gcod_planilla
    !num_control = Gcod_control
    !num_factura = Gcod_factura
    !codigo = txt_clientes.Text
    'El estatus del pedido es vigente
    'porque no se ha cancelado
    !status = "VI"
    !fecha_pedido = Me.txt_fecha_pedido.Text
    !id_inst = Me.dcmb_kgs.Text
    !monto_fac = CCur(Me.txt_precio.Text)
    !cant_pedido = Me.txt_cantidad.Text
    
    .Update

  With clientes.Recordset
    !fecha_ult_pago = Me.txt_fecha_pedido.Text
    !cedula = Me.TextBox4.Text
    .Update
  End With
End With

enlace1 = Gcod_planilla
enlace2 = Me.txt_clientes.Text

Me.cmd_ventas_materiales.Enabled = True
'Funcion que se encarga de cargar todos los pedidos
'de un cliente dado

Call buscar_pedidos

Call buscar_asignar_zona

With relacion.Recordset

    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
        .MoveLast
    End If
     .AddNew
       !usuario_liq = Usuario
       !id_pedido = Gcod_planilla
       !codigo = txt_clientes.Text
           
           If Me.dcmb_kgs.Text = "10" Then
            !descripcion = "Suministro GLP (10 Kgs)"
        End If
        
        If Me.dcmb_kgs.Text = "18" Then
            !descripcion = "Suministro GLP (18 Kgs)"
        End If
        
        If Me.dcmb_kgs.Text = "27" Then
            !descripcion = "Suministro GLP (27 Kgs)"
        End If
        
        If Me.dcmb_kgs.Text = "43" Then
            !descripcion = "Suministro GLP (43 Kgs)"
        End If
       
       !status = "VI"
       !fecha_pedido = Me.txt_fecha_pedido.Text
       !id_inst = Me.dcmb_kgs.Text
       !monto = CCur(Me.txt_precio.Text)
       !cant_pedido = Me.txt_cantidad.Text
       !marca = "1"
     .Update
End With

       
With facturando.Recordset

    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
        .MoveLast
    End If
    'Añadimos un pedido al cliente
    .AddNew
    
    !usuario_liq = Usuario
    !id_pedido = Gcod_planilla
    !num_control = Gcod_control
    !num_factura = Gcod_factura
    
    !codigo = txt_clientes.Text
    !fecha_pedido = Me.txt_fecha_pedido.Text
    !cliente = Me.txt_nombre.Text
    !cedula = Me.TextBox4.Text
    !direccion = Me.txt_direccion.Text
    !telefono_hab = Me.TextBox10.Text
    !observaciones = Me.txt_observaciones.Text
    !id_inst = Me.dcmb_kgs.Text
    
        If Me.dcmb_kgs.Text = "10" Then
            !descripcion = "Suministro GLP (10 Kgs)"
        End If
        
        If Me.dcmb_kgs.Text = "18" Then
            !descripcion = "Suministro GLP (18 Kgs)"
        End If
        
        If Me.dcmb_kgs.Text = "27" Then
            !descripcion = "Suministro GLP (27 Kgs)"
        End If
        
        If Me.dcmb_kgs.Text = "43" Then
            !descripcion = "Suministro GLP (43 Kgs)"
        End If
        
    !status = "VI"
    !marca = "1"
    !id_ruta = Me.txt_zona.Text
    !cant_pedido = Me.txt_cantidad.Text
    !monto_fac = CCur(Me.txt_precio.Text)
       total_iva = "0"
    !iva = total_iva
       total_factura = CCur(Me.txt_precio.Text + total_iva)
    !total_fac = total_factura
    
    .Update
End With

Me.txt_cantidad.Enabled = False

Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "JerGas")
        
    End Select

End Sub

Private Sub cmd_procesar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Me.cmd_procesar.FontBold = True

End Sub


'Private Sub cmd_liquidacion_Click()
'On Error GoTo control_error
''Desabilita el botón de aceptar
'Me.cmd_liquidacion.Enabled = False
'Screen.MousePointer = 11
'If DGrid_pedidos.SelBookmarks.Count = 0 Then
 
'    MsgBox "No se hallaron Pedidos marcados para Liquidar."
'    Me.cmd_liquidacion.Enabled = True
'    Screen.MousePointer = 0
'    Exit Sub

'End If
''Para cada registro seleccionado lo vamos a cancelar
''y generar su liquidación previa
'For Each Var In DGrid_pedidos.SelBookmarks
    
'    'Se crea la liquidación previa
'    'en la tabla liquidacion se coloca todo
'    'en estado vigente
'    DGrid_pedidos.Bookmark = Var
'    pedidos.Recordset.Bookmark = Var
    
'    With liquidado.Recordset
        
'        If Not (.BOF And .EOF) Then
        
'            mvBookMark = .Bookmark
                
'            .MoveLast
            
'        End If
        
'        'Añadimos la liquidacion del cliente
'        .AddNew
        
'        !id_pedido = DGrid_pedidos.Columns(0).Text
'        '!id_pedido = pedidos.Recordset!id_pedido
        
'        !usuario_liq = Usuario
        
'        !codigo = txt_clientes.Text
        'pedidos.Recordset!codigo
        
        'El estatus del pedido es cancelado
'        !Status = "CA"
        
'        !fecha_liq = Date
        
'        !monto = DGrid_pedidos.Columns(4).Text
        '!monto = pedidos.Recordset!monto_fac
        
'        .Update
    
'    End With
    
'    pedidos.Recordset.MoveFirst
    
'    strquery = "id_pedido = '" & DGrid_pedidos.Columns(0).Text & "'"

'    pedidos.Recordset.Find strquery
    
'    If pedidos.Recordset.EOF Then
    
'            MsgBox "Nºde Planilla suministrada no encontrada, por favor verifique ", vbInformation, "SIAGEP"
            
'            Screen.MousePointer = 0
                    
'    Else
    
'            pedidos.Recordset!Status = "CA"
            
'            pedidos.Recordset!fecha_cancel = Date
            
'            pedidos.Recordset.Update
        
'    End If
    
'Next

'rpt_liquidacion.Show

'hist_pedidos.Refresh

'Screen.MousePointer = 0

'Exit Sub

'control_error:
'Screen.MousePointer = 0
'    MsgBox Err.Description

'End Sub


Private Sub cmd_liquidacion_Click()
Dim actualiza As Integer
'On Error GoTo control_error

With resumen_mensual_ventas.Recordset
            
    mvBookMark = .Bookmark
    .MoveFirst
    .Find "fecha_venta =" & Me.txt_fecha.Text & ""
    
    If (.EOF) Then
     .MoveLast
     .AddNew
        !fecha_venta = Me.txt_fecha.Text
        !cant_10 = "0"
        !cant_18 = "0"
        !cant_27 = "0"
        !cant_43 = "0"
        !tot_10 = "0"
        !tot_18 = "0"
        !tot_27 = "0"
        !tot_43 = "0"
        
     .Update
    End If
End With

'Desabilita el botón de aceptar
Me.cmd_liquidacion.Enabled = False

Screen.MousePointer = 11

If DGrid_pedidos.SelBookmarks.Count = 0 Then
    
    MsgBox "No se hallaron Pedidos marcados para Liquidar."
    Me.cmd_liquidacion.Enabled = True
    Screen.MousePointer = 0
    Exit Sub

End If


'Para cada registro seleccionado lo vamos a cancelar
'y generar su liquidación previa
For Each Var In DGrid_pedidos.SelBookmarks
    
    'Se crea la liquidación previa
    'en la tabla liquidacion se coloca todo
    'en estado vigente
    DGrid_pedidos.Bookmark = Var
    
    Me.hist_pedidos.Recordset.Bookmark = Var
    Me.liquidado.Refresh
    
    With liquidado.Recordset
        
        If Not (.BOF And .EOF) Then
        
            mvBookMark = .Bookmark
                
            .MoveLast
            
        End If
        
        'Añadimos la liquidacion del cliente
        .AddNew
        
        !id_pedido = DGrid_pedidos.Columns(0).Text
        '!id_pedido = pedidos.Recordset!id_pedido
        
        !usuario_liq = Usuario
        
        !codigo = txt_clientes.Text
       ' !codigo = DGrid_pedidos.Columns(8).Text
        'pedidos.Recordset!codigo
        
        'El estatus del pedido es cancelado
        !status = "CA"
        
        !fecha_liq = Date
        
        !monto = CCur(DGrid_pedidos.Columns(4).Text)
        '!monto = pedidos.Recordset!monto_fac
        
        .Update
    
    End With
    
            hist_pedidos.Recordset!status = "CA"
            hist_pedidos.Recordset!fecha_cancel = Date
            hist_pedidos.Recordset.Update
        

Next

hist_pedidos.Refresh

Screen.MousePointer = 0

Exit Sub

control_error:
Screen.MousePointer = 0
    MsgBox Err.Description

End Sub

