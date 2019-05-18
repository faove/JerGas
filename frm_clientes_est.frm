VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_clientes_est 
   Caption         =   "SISTEMA DE GESTIÓN Y CONTROL JER-GAS, C.A."
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11640
   Icon            =   "frm_clientes_est.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   57
      Top             =   960
      Width           =   15735
      Begin VB.CommandButton Busquedad_avanzada 
         Caption         =   "Búsqueda Avanzada"
         Height          =   375
         Left            =   12840
         TabIndex        =   58
         Top             =   120
         Width           =   2175
      End
      Begin MSDataListLib.DataCombo Dcmb_Buscar 
         Bindings        =   "frm_clientes_est.frx":08CA
         Height          =   360
         Left            =   7560
         TabIndex        =   59
         ToolTipText     =   "Pulse doble click para cambiar el tipo de busqueda, después de presionar búsqueda avanzada"
         Top             =   120
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         ListField       =   "codigo"
         BoundColumn     =   "codigo"
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Búsqueda por Código:"
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
         Left            =   3960
         TabIndex        =   60
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.TextBox Text5 
      DataField       =   "id_ruta"
      DataSource      =   "rutas"
      Height          =   375
      Left            =   5760
      TabIndex        =   56
      Text            =   "Text5"
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text4 
      DataField       =   "status"
      DataSource      =   "status"
      Height          =   375
      Left            =   4920
      TabIndex        =   55
      Text            =   "Text4"
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Mostrar Datos del Propietario"
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
      Left            =   7200
      TabIndex        =   9
      Top             =   6000
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Left            =   1680
      Top             =   120
   End
   Begin VB.TextBox Text2 
      DataField       =   "fecha_ini_contrato"
      DataSource      =   "estantes"
      Height          =   375
      Left            =   3120
      TabIndex        =   53
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text3 
      DataField       =   "ncontrato"
      DataSource      =   "instalacion"
      Height          =   375
      Left            =   2280
      TabIndex        =   52
      TabStop         =   0   'False
      Text            =   "Text2"
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      DataField       =   "id"
      DataSource      =   "control_estantes"
      Height          =   375
      Left            =   4080
      TabIndex        =   51
      Text            =   "Text3"
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "|<<"
      Height          =   615
      Index           =   0
      Left            =   7200
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   6600
      Width           =   735
   End
   Begin VB.Frame Frame20 
      Height          =   1215
      Left            =   360
      TabIndex        =   44
      Top             =   6000
      Width           =   6495
      Begin VB.CommandButton cmdagregar 
         Caption         =   "&Agregar"
         Height          =   660
         Left            =   240
         TabIndex        =   49
         TabStop         =   0   'False
         ToolTipText     =   "Botón para Agregar un Nuevo Cliente"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdcancelar 
         Caption         =   "&Cancelar"
         Height          =   660
         Left            =   240
         TabIndex        =   48
         TabStop         =   0   'False
         ToolTipText     =   "Pulse este botón si desea Cancelar el Cliente Agregado"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdguardar 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   660
         Left            =   1440
         TabIndex        =   16
         ToolTipText     =   "Para Salvar el Cliente Agregado o Modificado"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdmodificar 
         Caption         =   "&Modificar"
         Height          =   660
         Left            =   2640
         TabIndex        =   47
         TabStop         =   0   'False
         ToolTipText     =   "Cambiar Característica de un Cliente"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdeliminar 
         Caption         =   "&Eliminar"
         Height          =   660
         Left            =   3840
         TabIndex        =   46
         TabStop         =   0   'False
         ToolTipText     =   "Elimina de la Base de Datos a un Cliente"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdsalir 
         Caption         =   "C&errar"
         Height          =   660
         Left            =   5040
         TabIndex        =   45
         TabStop         =   0   'False
         ToolTipText     =   "Cerrar y Volver al Menú Principal"
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<<"
      Height          =   615
      Index           =   1
      Left            =   7920
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      Height          =   615
      Index           =   2
      Left            =   8640
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>|"
      Height          =   615
      Index           =   3
      Left            =   9360
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   6600
      Width           =   735
   End
   Begin VB.Frame datos_pro 
      Caption         =   "DATOS DEL PROPIETARIO"
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
      Height          =   5775
      Left            =   10320
      TabIndex        =   19
      Top             =   1920
      Visible         =   0   'False
      Width           =   4695
      Begin VB.Frame Frame19 
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
         Left            =   240
         TabIndex        =   40
         ToolTipText     =   "Suministre los Teléfonos del Cliente."
         Top             =   3240
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
            DataSource      =   "estantes"
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
            TabIndex        =   13
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame18 
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
         TabIndex        =   39
         ToolTipText     =   "Suministre Apellido y Nombre del Cliente"
         Top             =   360
         Width           =   4335
         Begin VB.TextBox txt_propietario 
            DataField       =   "propietario"
            DataSource      =   "estantes"
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
            TabIndex        =   10
            Top             =   240
            Width           =   4095
         End
      End
      Begin VB.Frame Frame17 
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
         Left            =   240
         TabIndex        =   38
         ToolTipText     =   "Suministre la Cédula del Cliente."
         Top             =   1200
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
            DataSource      =   "estantes"
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
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame16 
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
         Height          =   1065
         Left            =   240
         TabIndex        =   37
         ToolTipText     =   "Suministre la Dirección del Cliente."
         Top             =   2040
         Width           =   4335
         Begin VB.TextBox txt_direccion_pro 
            DataField       =   "direccion_pro"
            DataSource      =   "estantes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   705
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   240
            Width           =   4095
         End
      End
      Begin VB.Frame Frame15 
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
         Left            =   240
         TabIndex        =   36
         ToolTipText     =   "Suministre el Correo Electrónico del Cliente."
         Top             =   4920
         Width           =   3855
         Begin VB.TextBox txt_correo 
            DataField       =   "e_mail"
            DataSource      =   "estantes"
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
            TabIndex        =   15
            Top             =   240
            Width           =   3615
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
         Left            =   240
         TabIndex        =   35
         ToolTipText     =   "Suministre los Teléfonos del Cliente."
         Top             =   4080
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
            DataSource      =   "estantes"
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
            TabIndex        =   14
            Top             =   240
            Width           =   2175
         End
      End
   End
   Begin MSAdodcLib.Adodc status 
      Height          =   450
      Left            =   240
      Top             =   8160
      Visible         =   0   'False
      Width           =   1965
      _ExtentX        =   3466
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
   Begin MSAdodcLib.Adodc control_estantes 
      Height          =   450
      Left            =   4560
      Top             =   7560
      Visible         =   0   'False
      Width           =   3045
      _ExtentX        =   5371
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
      RecordSource    =   "tbl_control_estantes"
      Caption         =   "control_estantes"
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
   Begin MSAdodcLib.Adodc estantes 
      Height          =   405
      Left            =   2280
      Top             =   8160
      Visible         =   0   'False
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   714
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
      Caption         =   "Estantes"
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
   Begin VB.Frame Frame1 
      Caption         =   "DATOS DE LA EMPRESA"
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
      Height          =   3855
      Left            =   240
      TabIndex        =   18
      Top             =   1920
      Width           =   9855
      Begin VB.Frame Frame21 
         Caption         =   "Fecha Ultimo Pedido"
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
         Height          =   705
         Left            =   7560
         TabIndex        =   54
         ToolTipText     =   "Suministre la Fecha del Contrato del Cliente."
         Top             =   1200
         Width           =   2055
         Begin MSComCtl2.DTPicker DTPicker2 
            Bindings        =   "frm_clientes_est.frx":08E2
            DataField       =   "fecha_ult_pago"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   3
            EndProperty
            DataSource      =   "estantes"
            Height          =   345
            Left            =   120
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
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
            Format          =   66125825
            CurrentDate     =   39083
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "NIF"
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
         Left            =   7320
         TabIndex        =   34
         Top             =   2880
         Width           =   2295
         Begin VB.TextBox txt_nif 
            Alignment       =   2  'Center
            DataField       =   "nif"
            DataSource      =   "estantes"
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
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "RIF"
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
         Left            =   4920
         TabIndex        =   33
         Top             =   2880
         Width           =   2295
         Begin VB.TextBox txt_rif 
            Alignment       =   2  'Center
            DataField       =   "rif"
            DataSource      =   "estantes"
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
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Teléfonos"
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
         TabIndex        =   32
         Top             =   2880
         Width           =   4575
         Begin VB.TextBox txt_telefono_emp2 
            Alignment       =   2  'Center
            DataField       =   "telefono_emp2"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "(####) ### ## ##"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "estantes"
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
            Left            =   2520
            TabIndex        =   6
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox txt_telefono_emp1 
            Alignment       =   2  'Center
            DataField       =   "telefono_emp1"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "(####) ### ## ##"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "estantes"
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
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   1935
         End
         Begin MSAdodcLib.Adodc rutas 
            Height          =   450
            Left            =   5040
            Top             =   0
            Visible         =   0   'False
            Width           =   1965
            _ExtentX        =   3466
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
         Begin MSAdodcLib.Adodc instalacion 
            Height          =   450
            Left            =   5760
            Top             =   360
            Visible         =   0   'False
            Width           =   2205
            _ExtentX        =   3889
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
      End
      Begin VB.Frame Frame4 
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
         TabIndex        =   30
         ToolTipText     =   "Suministre el Código del Cliente"
         Top             =   360
         Width           =   1815
         Begin VB.TextBox txt_codigo 
            Alignment       =   2  'Center
            DataField       =   "codigo"
            DataSource      =   "estantes"
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
            TabIndex        =   31
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
         TabIndex        =   28
         ToolTipText     =   "Suministre la Fecha del Contrato del Cliente."
         Top             =   360
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
            DataSource      =   "estantes"
            Height          =   345
            Left            =   240
            TabIndex        =   29
            TabStop         =   0   'False
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
            Format          =   66125825
            CurrentDate     =   39083
         End
      End
      Begin VB.Frame Frame33 
         Caption         =   "Nombre o Razón Social"
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
         TabIndex        =   27
         ToolTipText     =   "Suministre Apellido y Nombre del Cliente"
         Top             =   1200
         Width           =   7215
         Begin VB.TextBox txt_cliente 
            DataField       =   "cliente"
            DataSource      =   "estantes"
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
            TabIndex        =   2
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
         TabIndex        =   25
         ToolTipText     =   "Suministre el Nº de Contrato del Cliente"
         Top             =   360
         Width           =   1575
         Begin VB.TextBox txt_contrato 
            Alignment       =   2  'Center
            DataField       =   "contrato_num"
            DataSource      =   "estantes"
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
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   240
            Width           =   1215
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
         TabIndex        =   24
         ToolTipText     =   "Suministre la Dirección del Cliente."
         Top             =   2040
         Width           =   9375
         Begin VB.TextBox txt_direccion 
            DataField       =   "direccion"
            DataSource      =   "estantes"
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
            Width           =   9135
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
         TabIndex        =   23
         ToolTipText     =   "Suministre el Tipo de Contrato del Cliente."
         Top             =   360
         Width           =   1095
         Begin MSDataListLib.DataCombo txt_cilindro 
            Bindings        =   "frm_clientes_est.frx":08FD
            DataField       =   "id_inst"
            DataSource      =   "estantes"
            Height          =   315
            Left            =   120
            TabIndex        =   0
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
         TabIndex        =   21
         ToolTipText     =   "Suministre el Status del Cliente."
         Top             =   360
         Width           =   1095
         Begin MSDataListLib.DataCombo txt_status 
            Bindings        =   "frm_clientes_est.frx":0917
            DataField       =   "status"
            DataSource      =   "estantes"
            Height          =   315
            Left            =   120
            TabIndex        =   22
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
         TabIndex        =   20
         ToolTipText     =   "Suministre el Tipo de Contrato del Cliente."
         Top             =   360
         Width           =   1095
         Begin MSDataListLib.DataCombo txt_ruta 
            Bindings        =   "frm_clientes_est.frx":092C
            DataField       =   "id_ruta"
            DataSource      =   "estantes"
            Height          =   315
            Left            =   120
            TabIndex        =   1
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
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   450
      Left            =   240
      Top             =   7560
      Visible         =   0   'False
      Width           =   1965
      _ExtentX        =   3466
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   450
      Left            =   2280
      Top             =   7560
      Visible         =   0   'False
      Width           =   2205
      _ExtentX        =   3889
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
   Begin VB.Shape Shape2 
      BackColor       =   &H80000000&
      BorderColor     =   &H8000000B&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   330
      Left            =   0
      Top             =   900
      Width           =   15465
   End
   Begin VB.Label Label_titulo 
      BackColor       =   &H80000001&
      Caption         =   "  CONTROL DE CLIENTES (ESTANTES)"
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
      Left            =   6600
      TabIndex        =   17
      Top             =   240
      Width           =   8655
   End
End
Attribute VB_Name = "frm_clientes_est"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Busquedad_avanzadas_Click(Index As Integer)
            Busq_Avanzada = True
            
            estantes.CommandType = adCmdText
            
            estantes.RecordSource = "select * from tbl_estantes WHERE codigo <> '' ORDER BY codigo ASC"
            
            estantes.Refresh
            
            Call Dcmb_Buscar_Click(1)
End Sub

Private Sub Check1_Click()

If Check1 = 1 Then
 datos_pro.Visible = True
Else
 datos_pro.Visible = False
End If
End Sub

Private Sub Dcmb_Buscar_Click(Area As Integer)
If Area = 2 Then
        If Dcmb_Buscar.ListField = "codigo" Then
            If Dcmb_Buscar.Text <> "" Then
                
                Call buscar_codigo
            Else
                Exit Sub
            End If
        End If
        
        If Dcmb_Buscar.ListField = "contrato_num" Then
            If Dcmb_Buscar.Text <> "" Then
                Call buscar_contrato_num
            Else
                Exit Sub
            End If
        End If

        If Dcmb_Buscar.ListField = "cliente" Then
            If Dcmb_Buscar.Text <> "" Then
                Call buscar_cliente
            Else
                Exit Sub
            End If
        End If

        If Dcmb_Buscar.ListField = "cedula" Then
            If Dcmb_Buscar.Text <> "" Then
                Call buscar_cedula
              Else
                Exit Sub
            End If
        End If
        If Dcmb_Buscar.ListField = "direccion" Then
            If Dcmb_Buscar.Text <> "" Then
                Call buscar_direccion
            Else
                Exit Sub
            End If
        End If
End If
End Sub

Private Sub cmdagregar_Click()
On Error GoTo AddErr
    DTPicker1.Value = Date
    
    Dim fec As Date
    Dim ano As Date
    
    cmdagregar.Visible = False
    cmdeliminar.Enabled = False
    cmdmodificar.Enabled = False
    cmdcancelar.Enabled = True
    cmdsalir.Enabled = False
    
    cmdguardar.Enabled = True
    cmdcancelar.Visible = True
    
    Me.Command1(0).Enabled = False
    Me.Command1(1).Enabled = False
    Me.Command1(2).Enabled = False
    Me.Command1(3).Enabled = False
    
    DTPicker1.Enabled = True
    txt_ruta.Locked = False
    txt_status.Locked = False
    txt_cilindro.Locked = False
    txt_cliente.Locked = False
    txt_direccion.Locked = False
    txt_telefono_emp1.Locked = False
    txt_telefono_emp2.Locked = False
    txt_rif.Locked = False
    txt_nif.Locked = False
    
    txt_propietario.Locked = False
    txt_cedula.Locked = False
    txt_direccion_pro.Locked = False
    txt_telefono_hab.Locked = False
    txt_telefono_cel.Locked = False
    txt_correo.Locked = False
    
    fecha = DTPicker1.Value
    
    Text1.Text = DateAdd("m", 1, fec)
    
    Text2.Text = fecha
    
    Text2.Text = DateAdd("yyyy", 1, ano)
    
    With estantes.Recordset
    
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
            
        .MoveLast
    End If
    .AddNew
      
  End With
    Me.DTPicker1.Value = Date
    Me.txt_status.Text = "VI"
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdcancelar_Click()

On Error GoTo ControlError
    
    estantes.Recordset.CancelUpdate
'    If mvBookMark > 0 Then
'        estantes.Recordset.Bookmark = mvBookMark
'    Else
'        estantes.Recordset.MoveFirst
'    End If
'
'       estantes.Recordset.CancelUpdate
            cmdguardar.Enabled = False
            cmdagregar.Visible = True
            cmdeliminar.Enabled = True
            cmdmodificar.Enabled = True
            cmdsalir.Enabled = True
            cmdcancelar.Visible = False
            Me.Command1(0).Enabled = True
            Me.Command1(1).Enabled = True
            Me.Command1(2).Enabled = True
            Me.Command1(3).Enabled = True
            
                txt_ruta.Locked = True
                txt_status.Locked = True
                txt_cilindro.Locked = True
                txt_cliente.Locked = True
                txt_direccion.Locked = True
                txt_telefono_emp1.Locked = True
                txt_telefono_emp2.Locked = True
                txt_rif.Locked = True
                txt_nif.Locked = True
    
                txt_propietario.Locked = True
                txt_cedula.Locked = True
                txt_direccion_pro.Locked = True
                txt_telefono_hab.Locked = True
                txt_telefono_cel.Locked = True
                txt_correo.Locked = True

    Exit Sub    ' Salir para evitar el controlador.

ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "Control de Cliente")
        
    End Select

 End Sub

Private Sub cmdeliminar_Click()
  'esto puede producir un error si elimina el último
  'registro o el único registro del recordset

On Error GoTo ControlError
    respuesta = MsgBox("¿Desea Eliminar el Registro?", vbYesNo)
    If respuesta = vbYes Then
        estantes.Recordset.Delete
        estantes.Recordset.MoveNext
    End If

    Exit Sub    ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "Control de Clientes")
    End Select
End Sub

Private Sub cmdmodificar_Click()
    
    cmdagregar.Visible = False
    cmdeliminar.Enabled = False
    cmdcancelar.Visible = True
    cmdsalir.Enabled = False
    cmdguardar.Enabled = True
    cmdmodificar.Enabled = False
 
       Command1(0).Enabled = False
       Command1(1).Enabled = False
       Command1(2).Enabled = False
       Command1(3).Enabled = False
        
    txt_ruta.Locked = False
    txt_status.Locked = False
    txt_cilindro.Locked = False
    txt_cliente.Locked = False
    txt_direccion.Locked = False
    txt_telefono_emp1.Locked = False
    txt_telefono_emp2.Locked = False
    txt_rif.Locked = False
    txt_nif.Locked = False
    
    txt_propietario.Locked = False
    txt_cedula.Locked = False
    txt_direccion_pro.Locked = False
    txt_telefono_hab.Locked = False
    txt_telefono_cel.Locked = False
    txt_correo.Locked = False


End Sub

Private Sub cmdguardar_Click()
Dim fec As Date
Dim ano As Date
Dim strquery As String
Dim bandera As Boolean
Dim abc, ncliente, contrato

On Error GoTo ControlError
    
    
    If IsNull(txt_cliente.Text) Or txt_cliente.Text = "" Then
    
        MsgBox "El Nombre no puede ser nulo, por favor verifique ", vbInformation, "JerGas C.A."
             
        Me.txt_cliente.SetFocus
             
        Exit Sub
    End If
    
    If IsNull(Me.txt_ruta.Text) Or txt_ruta.Text = "" Then
    
        MsgBox "La ruta no puede ser nulo, por favor verifique ", vbInformation, "JerGas C.A."
             
        Me.txt_ruta.SetFocus
             
        Exit Sub
    End If
    
    If IsNull(Me.txt_cilindro.Text) Or txt_cilindro.Text = "" Then
    
        MsgBox "Debe seleccionar un cilindro, por favor verifique ", vbInformation, "JerGas C.A."
             
        Me.txt_cilindro.SetFocus
             
        Exit Sub
    End If
    
    'Buscamos el ultimo numero de cliente
    
     control_estantes.Recordset.MoveFirst
    
    abc = Left(txt_cliente.Text, 1)
    
    strquery = "id = '" & abc & "'"
    
    control_estantes.Recordset.Find strquery
    
    If control_estantes.Recordset.EOF Then
    
             MsgBox "En el campo Apellido y Nombre la primera caracter suministrado debe ser una letra, por favor verifique ", vbInformation, "JerGas"
             
             Me.txt_cliente.SetFocus
             
             Exit Sub
    Else
        ncliente = CInt(control_estantes.Recordset!valor) + 1
    
        ncliente = Format(ncliente, "0000")
    
        txt_codigo.Text = "" + abc + "-" + Me.txt_ruta.Text + "-" + ncliente + ""
        
        control_estantes.Recordset!valor = ncliente
        
        control_estantes.Recordset.Update
        
         
    End If
      
    'Buscamos la instalacion con respecto a la bonbona seleccionada
    
     'Buscamos el ultimo numero de cliente
    
    instalacion.Recordset.MoveFirst
    
    strquery = "id_inst = '" & Me.txt_cilindro.Text & "'"
    
    instalacion.Recordset.Find strquery
    
    If instalacion.Recordset.EOF Then
    
             MsgBox "Verifique la bombona suministrada ", vbInformation, "JerGas C.A."
             
             Me.txt_cilindro.SetFocus
             
             Exit Sub
    Else
    
        contrato = CInt(instalacion.Recordset!ncontrato) + 1
    
        contrato = Format(contrato, "00000")
    
        Me.txt_contrato.Text = "" + Me.txt_cilindro.Text + "-" + contrato + ""
        
        instalacion.Recordset!ncontrato = contrato
        
        instalacion.Recordset.Update
    
    End If
   
    With estantes.Recordset

        mvBookMark = .Bookmark

        .Update

        .Bookmark = mvBookMark
      ' estantes.Refresh

    End With
    
    cmdeliminar.Enabled = True
    cmdmodificar.Enabled = True
    cmdagregar.Visible = True
    cmdsalir.Enabled = True
    cmdcancelar.Visible = False
    cmdguardar.Enabled = False
   
       Command1(0).Enabled = True
       Command1(1).Enabled = True
       Command1(2).Enabled = True
       Command1(3).Enabled = True

    cmdagregar.SetFocus
    
    Exit Sub    ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "Control de Clientes")
        Case 3314
            MsgBox "Verifique la Cédula ", vbOKOnly, "Control de Clientes"
        Case 524
            MsgBox "Verifique el Nombre y Apellido  ", vbOKOnly, "Control de Clientes"
        Case -2147467259
            MsgBox "Error, la cédula suministrada ya existe", vbOKOnly, "Control de Clientes"
        Case -2147217842
            MsgBox "Error, cancele la operación y vuelva a intentarlo", vbOKOnly, "Control de Clientes"
        Case -2147217887
            MsgBox "Error, al crear histórico, se recomienda borrar el registro y volverlo a crear", vbOKOnly, "Control de Clientes"
    End Select
End Sub

Private Sub Form_Load()
'Call actualizar_cn("SQL Server")
    cmdcancelar.Visible = False
    DTPicker1.Enabled = True
    'estantes.Refresh
    Text1.Text = Date
Me.Command1(0).Enabled = True
Me.Command1(1).Enabled = True
Me.Command1(2).Enabled = True
Me.Command1(3).Enabled = True

End Sub

Private Sub cmdsalir_Click()
  Unload Me
End Sub

Private Sub Form_Resize()
'Call Mover_der(Me, Label_titulo, 0)
'Call Mover_centrado(Me, Frame1)
'Call Mover_der(Me, Frame3, 10)
'Call Mover_der(Me, Me.LABEL_BUSCA, Frame3.Width + 15)
'Shape1.Width = Me.Width
'Shape1.Left = 0
End Sub

Private Sub Timer1_Timer()
'Esta función se encarga de sumar un mes
' y sumar un año para establecer en proximo pago del cliente
'Dim fec As Date
'Dim ano As Date
'
'    fec = DTPicker1.Value
'    txtFields(9).Text = DateAdd("m", 1, fec)
'    ano = DTPicker2.Value
'    txtFields(4).Text = DateAdd("yyyy", 1, ano)
   
End Sub

Private Sub Command1_Click(Index As Integer)
On Error GoTo ControlError

Select Case Index
    Case 0
       estantes.Recordset.MoveFirst
       Command1(0).Enabled = False
       Command1(1).Enabled = False
       Command1(2).Enabled = True
       Command1(3).Enabled = True
         Me.cmdsalir.FontBold = False
         Me.cmdguardar.FontBold = False
         Me.cmdeliminar.FontBold = False
         Me.cmdcancelar.FontBold = False
         Me.cmdmodificar.FontBold = False
         Me.cmdagregar.FontBold = False
    Case 1
       estantes.Recordset.MovePrevious
       Command1(2).Enabled = True
       Command1(3).Enabled = True
         Me.cmdsalir.FontBold = False
         Me.cmdguardar.FontBold = False
         Me.cmdeliminar.FontBold = False
         Me.cmdcancelar.FontBold = False
         Me.cmdmodificar.FontBold = False
         Me.cmdagregar.FontBold = False
       If estantes.Recordset.BOF = True Then
        estantes.Recordset.MoveFirst
        Command1(0).Enabled = False
        Command1(1).Enabled = False
         Me.cmdsalir.FontBold = False
         Me.cmdguardar.FontBold = False
         Me.cmdeliminar.FontBold = False
         Me.cmdcancelar.FontBold = False
         Me.cmdmodificar.FontBold = False
         Me.cmdagregar.FontBold = False
       Else
        End If
    Case 2
       estantes.Recordset.MoveNext
       Command1(0).Enabled = True
       Command1(1).Enabled = True
         Me.cmdsalir.FontBold = False
         Me.cmdguardar.FontBold = False
         Me.cmdeliminar.FontBold = False
         Me.cmdcancelar.FontBold = False
         Me.cmdmodificar.FontBold = False
         Me.cmdagregar.FontBold = False
       If estantes.Recordset.EOF = True Then
         Command1(2).Enabled = False
         Command1(3).Enabled = False
         Me.cmdsalir.FontBold = False
         Me.cmdguardar.FontBold = False
         Me.cmdeliminar.FontBold = False
         Me.cmdcancelar.FontBold = False
         Me.cmdmodificar.FontBold = False
         Me.cmdagregar.FontBold = False
         estantes.Recordset.MoveLast
       Else
       End If
    Case 3
       estantes.Recordset.MoveLast
       Command1(0).Enabled = True
       Command1(1).Enabled = True
       Command1(2).Enabled = False
       Command1(3).Enabled = False
         Me.cmdsalir.FontBold = False
         Me.cmdguardar.FontBold = False
         Me.cmdeliminar.FontBold = False
         Me.cmdcancelar.FontBold = False
         Me.cmdmodificar.FontBold = False
         Me.cmdagregar.FontBold = False
End Select

ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
'        Case 2021
'            v = MsgBox("Formato No Válido", vbOKOnly, "Control del Cliente")
'        Case 3314
'            MsgBox "Verifique la Cédula ", vbOKOnly, "Control del Cliente"
'        Case 524
'            MsgBox "Verifique el Nombre y Apellido  ", vbOKOnly, "Control del Cliente"
        Case -2147467259
            MsgBox "Error, la cédula suministrada ya existe", vbOKOnly, "Control de Clientes"
'        Case -2147217842
'            MsgBox "Error, cancele la operación y vuelva a intentarlo", vbOKOnly, "Control del Cliente"
'        Case -2147217887
'            MsgBox "Error, al crear histórico, se recomienda borrar el registro y volverlo a crear", vbOKOnly, "Control del Cliente"
    End Select
End Sub

Private Sub cmdmodificar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdmodificar.FontBold = True
Me.cmdagregar.FontBold = False
Me.cmdguardar.FontBold = False
Me.cmdeliminar.FontBold = False
Me.cmdcancelar.FontBold = False
Me.cmdsalir.FontBold = False
End Sub

Private Sub cmdguardar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdguardar.FontBold = True
Me.cmdagregar.FontBold = False
Me.cmdmodificar.FontBold = False
Me.cmdeliminar.FontBold = False
Me.cmdcancelar.FontBold = False
Me.cmdsalir.FontBold = False


End Sub

Private Sub cmdeliminar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdeliminar.FontBold = True
Me.cmdagregar.FontBold = False
Me.cmdguardar.FontBold = False
Me.cmdmodificar.FontBold = False
Me.cmdcancelar.FontBold = False
Me.cmdsalir.FontBold = False
End Sub

Private Sub cmdcancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdcancelar.FontBold = True
Me.cmdagregar.FontBold = False
Me.cmdguardar.FontBold = False
Me.cmdeliminar.FontBold = False
Me.cmdmodificar.FontBold = False
Me.cmdsalir.FontBold = False
End Sub

Private Sub cmdagregar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdagregar.FontBold = True
Me.cmdguardar.FontBold = False
Me.cmdeliminar.FontBold = False
Me.cmdcancelar.FontBold = False
Me.cmdmodificar.FontBold = False
Me.cmdsalir.FontBold = False
End Sub

Private Sub cmdsalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdsalir.FontBold = True
Me.cmdguardar.FontBold = False
Me.cmdeliminar.FontBold = False
Me.cmdcancelar.FontBold = False
Me.cmdmodificar.FontBold = False
Me.cmdagregar.FontBold = False
End Sub

Private Sub Dcmb_Buscar_DblClick(Area As Integer)
'Esta funci[on redefine el tipo de busqueda
'If Busq_Avanzada Then

    Me.Dcmb_Buscar.ToolTipText = "Doble click para alternar el tipo de busqueda"
    
    If Dcmb_Buscar.ListField = "codigo" Then
        'Si es bif pasa a ape
        If Busq_Avanzada Then
            Me.estantes.CommandType = adCmdText
            Me.estantes.RecordSource = "select * from tbl_estantes WHERE tbl_estantes.contrato_num <> '' ORDER BY contrato_num ASC"
            Me.estantes.Refresh
        End If

        Dcmb_Buscar.ListField = "contrato_num"
        Dcmb_Buscar.Text = ""
        LABEL_BUSCA.Caption = "Búsqueda por Nº de Contrato:"
        Exit Sub
    End If

    If Dcmb_Buscar.ListField = "contrato_num" Then
    
        'Si es ape pasa a cod cata
        If Busq_Avanzada Then
            Me.estantes.CommandType = adCmdText
            Me.estantes.RecordSource = "select * from tbl_estantes WHERE tbl_estantes.cliente <> '' ORDER BY cliente ASC"
            Me.estantes.Refresh
        End If
        
        Dcmb_Buscar.ListField = "cliente"
        Dcmb_Buscar.Text = ""
        LABEL_BUSCA.Caption = "Búsqueda por Cliente:"
        Exit Sub
        
    End If

    If Dcmb_Buscar.ListField = "cliente" Then

        'Si es cod pasa a cedula
        If Busq_Avanzada Then
            estantes.CommandType = adCmdText
            estantes.RecordSource = "select * from tbl_estantes WHERE tbl_estantes.cedula <> '' ORDER BY cedula ASC"
            estantes.Refresh
        End If
        
        Dcmb_Buscar.ListField = "cedula"
        Dcmb_Buscar.Text = ""
        LABEL_BUSCA.Caption = "Búsqueda por Cédula: "
        Exit Sub

    End If

    If Dcmb_Buscar.ListField = "cedula" Then
    
        'Si es cedual pasa a direccion
        If Busq_Avanzada Then
            estantes.CommandType = adCmdText
            estantes.RecordSource = "select * from tbl_estantes WHERE tbl_estantes.direccion <> '' ORDER BY direccion ASC"
            estantes.Refresh
        End If
        
        Dcmb_Buscar.ListField = "direccion"
        Dcmb_Buscar.Text = ""
        LABEL_BUSCA.Caption = "Búsqueda por Dirección: "
        Exit Sub
    End If

    If Dcmb_Buscar.ListField = "direccion" Then

        'Si es direccion pasa a bif
        If Busq_Avanzada Then
            estantes.CommandType = adCmdText
            estantes.RecordSource = "select * from tbl_estantes WHERE codigo <> '' ORDER BY codigo ASC"
            estantes.Refresh
        End If
        
        Dcmb_Buscar.ListField = "codigo"
        Dcmb_Buscar.Text = ""
        LABEL_BUSCA.Caption = "Búsqueda por Código: "
        Exit Sub
    End If
'End If
End Sub

Private Sub buscar_codigo()

On Error GoTo ControlError
Dim strquery, RESP

If Not Busq_Avanzada And ((Dcmb_Buscar.Text Like "%*%" Or Dcmb_Buscar.Text Like "%*" Or Dcmb_Buscar.Text Like "*%") Or (Me.estantes.Recordset.RecordCount = 0)) Then
    
    Me.estantes.CommandType = adCmdText
    Me.estantes.RecordSource = "SELECT * FROM tbl_estantes WHERE tbl_estantes.codigo like '" & Dcmb_Buscar.Text & "' ORDER BY codigo"
    Me.estantes.Refresh
    
    If estantes.Recordset.EOF Then
        MsgBox "Código suministrado no encontrado, por favor verifique ", vbInformation, "JerGas C.A."
        Dcmb_Buscar.Text = ""
        Dcmb_Buscar.SetFocus
        Call habilitar_botones(False)
    Else
        If Me.estantes.Recordset.RecordCount > 1 Then
            MsgBox Me.estantes.Recordset.RecordCount & " encontrados"
            Busq_Avanzada = True
        End If
       
        Call habilitar_botones(True)
    End If
Else
    estantes.Recordset.MoveFirst
    strquery = "codigo = '" & Dcmb_Buscar.Text & "'"
    estantes.Recordset.Find strquery
    
    If estantes.Recordset.EOF Then
             MsgBox "Código suministrado no encontrado, por favor verifique ", vbInformation, "JerGas C.A."
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
            v = MsgBox("Formato No Válido", vbOKOnly, "JerGas C.A.")
        Case 3001
            v = MsgBox("Nombre suministrado no encontrado", vbOKOnly, "JerGas C.A.")
    End Select

End Sub

Private Sub buscar_contrato_num()

On Error GoTo ControlError

Dim strquery, RESP

If Not Busq_Avanzada And ((Dcmb_Buscar.Text Like "%*%" Or Dcmb_Buscar.Text Like "%*" Or Dcmb_Buscar.Text Like "*%") Or (Me.estantes.Recordset.RecordCount = 0)) Then
    
    Me.estantes.CommandType = adCmdText
    
    Me.estantes.RecordSource = "SELECT * FROM tbl_estantes WHERE tbl_estantes.contrato_num like '" & Dcmb_Buscar.Text & "' ORDER BY contrato_num"
    
    Me.estantes.Refresh

    If estantes.Recordset.EOF Then
    
        MsgBox "Nº de Contrato suministrado no encontrado, por favor verifique ", vbInformation, "JerGas C.A."
        
        Dcmb_Buscar.Text = ""
        
        Dcmb_Buscar.SetFocus
        
        Call habilitar_botones(False)
        
    Else
    
        If Me.estantes.Recordset.RecordCount > 1 Then
        
            MsgBox Me.estantes.Recordset.RecordCount & " encontrados"
            
            Busq_Avanzada = True
            
        End If
        
        Call habilitar_botones(True)
        
    End If
    
Else
    
    estantes.Recordset.MoveFirst
    
    strquery = "contrato_num = '" & Dcmb_Buscar.Text & "'"

    estantes.Recordset.Find strquery
    
    If estantes.Recordset.EOF Then
    
            MsgBox "Nº de Contrato suministrado no encontrado, por favor verifique ", vbInformation, "JerGas C.A."
            
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
            v = MsgBox("Formato No Válido", vbOKOnly, "JerGas C.A.")
        Case 3001
            v = MsgBox("Contrato suministrado no encontrado", vbOKOnly, "JerGas C.A.")
    End Select

End Sub
Private Sub buscar_cedula()

On Error GoTo ControlError

Dim strquery, RESP

If Not Busq_Avanzada And ((Dcmb_Buscar.Text Like "%*%" Or Dcmb_Buscar.Text Like "%*" Or Dcmb_Buscar.Text Like "*%") Or (Me.estantes.Recordset.RecordCount = 0)) Then
    
    Me.estantes.CommandType = adCmdText
    
    Me.estantes.RecordSource = "SELECT * FROM tbl_estantes WHERE tbl_estantes.cedula like '" & Dcmb_Buscar.Text & "' ORDER BY cedula"
    
    Me.estantes.Refresh

    If estantes.Recordset.EOF Then
    
        MsgBox "Cédula del cliente suministrado no encontrado, por favor verifique ", vbInformation, "JerGas C.A."
        
        Dcmb_Buscar.Text = ""
        
        Dcmb_Buscar.SetFocus
        
        Call habilitar_botones(False)
        
    Else
    
        If Me.estantes.Recordset.RecordCount > 1 Then
        
            MsgBox Me.estantes.Recordset.RecordCount & " encontrados"
            
            Busq_Avanzada = True
            
        End If
        
        Call habilitar_botones(True)
        
    End If
    
Else
    
    estantes.Recordset.MoveFirst
    
    strquery = "cedula = '" & Dcmb_Buscar.Text & "'"

    estantes.Recordset.Find strquery
    
    If estantes.Recordset.EOF Then
    
            MsgBox "Cédula del Cliente suministrado no encontrado, por favor verifique ", vbInformation, "JerGas C.A."
            
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
            v = MsgBox("Formato No Válido", vbOKOnly, "JerGas C.A.")
        Case 3001
            v = MsgBox("Cédula del Cliente suministrado no encontrado", vbOKOnly, "JerGas C.A.")
    End Select

End Sub

Private Sub buscar_cliente()

On Error GoTo ControlError

Dim strquery, RESP

If Not Busq_Avanzada And ((Dcmb_Buscar.Text Like "%*%" Or Dcmb_Buscar.Text Like "%*" Or Dcmb_Buscar.Text Like "*%") Or (Me.estantes.Recordset.RecordCount = 0)) Then
    
    Me.estantes.CommandType = adCmdText
    
    Me.estantes.RecordSource = "SELECT * FROM tbl_estantes WHERE tbl_estantes.cliente like '" & Dcmb_Buscar.Text & "' ORDER BY cliente"
    
    Me.estantes.Refresh

    If estantes.Recordset.EOF Then
    
        MsgBox "Nombre del cliente suministrado no encontrado, por favor verifique ", vbInformation, "JerGas C.A."
        
        Dcmb_Buscar.Text = ""
        
        Dcmb_Buscar.SetFocus
        
        Call habilitar_botones(False)
        
    Else
    
        If Me.estantes.Recordset.RecordCount > 1 Then
        
            MsgBox Me.estantes.Recordset.RecordCount & " encontrados"
            
            Busq_Avanzada = True
            
        End If
        
        Call habilitar_botones(True)
        
    End If
    
Else
    
    estantes.Recordset.MoveFirst
    
    strquery = "cliente = '" & Dcmb_Buscar.Text & "'"

    estantes.Recordset.Find strquery
    
    If estantes.Recordset.EOF Then
    
            MsgBox "Nombre del Cliente suministrado no encontrado, por favor verifique ", vbInformation, "JerGas C.A."
            
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
            v = MsgBox("Formato No Válido", vbOKOnly, "JerGas C.A.")
        Case 3001
            v = MsgBox("Nombre del Cliente suministrado no encontrado", vbOKOnly, "JerGas C.A.")
    End Select

End Sub

Private Sub buscar_direccion()

On Error GoTo ControlError

Dim strquery, RESP

If Not Busq_Avanzada And ((Dcmb_Buscar.Text Like "%*%" Or Dcmb_Buscar.Text Like "%*" Or Dcmb_Buscar.Text Like "*%") Or (Me.estantes.Recordset.RecordCount = 0)) Then
    
    Me.estantes.CommandType = adCmdText
    
    Me.estantes.RecordSource = "SELECT * FROM tbl_estantes WHERE tbl_estantes.direccion like '" & Dcmb_Buscar.Text & "' ORDER BY direccion"
    
    Me.estantes.Refresh

    If estantes.Recordset.EOF Then
    
        MsgBox "Direccion del cliente suministrada no encontrada, por favor verifique ", vbInformation, "JerGas C.A."
        
        Dcmb_Buscar.Text = ""
        
        Dcmb_Buscar.SetFocus
        
        Call habilitar_botones(False)
        
    Else
    
        If Me.estantes.Recordset.RecordCount > 1 Then
        
            MsgBox Me.estantes.Recordset.RecordCount & " encontrados"
            
            Busq_Avanzada = True
            
        End If
        
        Call habilitar_botones(True)
        
    End If
    
Else
    
    estantes.Recordset.MoveFirst
    
    strquery = "direccion = '" & Dcmb_Buscar.Text & "'"

    estantes.Recordset.Find strquery
    
    If estantes.Recordset.EOF Then
    
            MsgBox "Dirección del Cliente suministrado no encontrado, por favor verifique ", vbInformation, "JerGas C.A."
            
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
            v = MsgBox("Formato No Válido", vbOKOnly, "JerGas C.A.")
        Case 3001
            v = MsgBox("Dirección del Cliente suministrado no encontrado", vbOKOnly, "JerGas C.A.")
    End Select

End Sub


Private Sub habilitar_botones(valor As Boolean)
'Me.cmd_estantes.Enabled = valor
'Me.cmd_estado.Enabled = valor
'Me.cmd_liquidacion.Enabled = valor

End Sub

Private Sub txt_ruta_KeyPress(KeyAscii As Integer)
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

Private Sub txt_telefono_emp1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
   ' End If
  End Sub

Private Sub txt_telefono_emp2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
 '   End If
End Sub

Private Sub txt_rif_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Index <> 6 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 13 Or Index = 6 Or Index = 9 Then
        If ((KeyAscii < 48) Or (KeyAscii > 57)) And (KeyAscii <> 8) Then KeyAscii = 0
    End If
End Sub

Private Sub txt_nif_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Index <> 6 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 13 Or Index = 6 Or Index = 9 Then
        If ((KeyAscii < 48) Or (KeyAscii > 57)) And (KeyAscii <> 8) Then KeyAscii = 0
    End If
End Sub

Private Sub txt_propietario_KeyPress(KeyAscii As Integer)
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
 ' End If
End Sub

Private Sub txt_direccion_pro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Index <> 6 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 13 Or Index = 6 Or Index = 9 Then
        If ((KeyAscii < 48) Or (KeyAscii > 57)) And (KeyAscii <> 8) Then KeyAscii = 0
    End If
End Sub

Private Sub txt_telefono_hab_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
   ' End If
  End Sub

Private Sub txt_telefono_cel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
 '   End If
End Sub

Private Sub txt_correo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Index <> 6 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 13 Or Index = 6 Or Index = 9 Then
        If ((KeyAscii < 48) Or (KeyAscii > 57)) And (KeyAscii <> 8) Then KeyAscii = 0
    End If
End Sub

