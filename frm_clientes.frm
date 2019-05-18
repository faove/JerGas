VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_clientes 
   Caption         =   "SISTEMA DE GESTIÓN Y CONTROL JER-GAS, C.A."
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11700
   Icon            =   "frm_clientes.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8040
   ScaleWidth      =   11700
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text3 
      DataField       =   "id"
      DataSource      =   "control_clientes"
      Height          =   375
      Left            =   1320
      TabIndex        =   29
      Text            =   "Text3"
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   480
      TabIndex        =   22
      Top             =   1800
      Width           =   9495
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
         Left            =   8280
         TabIndex        =   45
         ToolTipText     =   "Suministre el Tipo de Contrato del Cliente."
         Top             =   0
         Width           =   1095
         Begin MSDataListLib.DataCombo txt_ruta 
            Bindings        =   "frm_clientes.frx":08CA
            DataField       =   "id_ruta"
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
         Left            =   5880
         TabIndex        =   43
         ToolTipText     =   "Suministre el Status del Cliente."
         Top             =   0
         Width           =   1095
         Begin MSDataListLib.DataCombo txt_status 
            Bindings        =   "frm_clientes.frx":08DE
            DataField       =   "status"
            DataSource      =   "clientes"
            Height          =   315
            Left            =   120
            TabIndex        =   44
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
         Left            =   7080
         TabIndex        =   42
         ToolTipText     =   "Suministre el Tipo de Contrato del Cliente."
         Top             =   0
         Width           =   1095
         Begin MSDataListLib.DataCombo txt_cilindro 
            Bindings        =   "frm_clientes.frx":08F3
            DataField       =   "id_inst"
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
         Left            =   2760
         TabIndex        =   41
         ToolTipText     =   "Suministre los Teléfonos del Cliente."
         Top             =   2520
         Width           =   2415
         Begin VB.TextBox txt_telefono_cel 
            Alignment       =   2  'Center
            DataField       =   "telefono_cel"
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
         Left            =   0
         TabIndex        =   40
         ToolTipText     =   "Suministre Cualquier Dato Adicional del Cliente."
         Top             =   3360
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
            TabIndex        =   9
            Top             =   240
            Width           =   5175
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
         Left            =   5520
         TabIndex        =   39
         ToolTipText     =   "Suministre el Correo Electrónico del Cliente."
         Top             =   2520
         Width           =   3855
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
            TabIndex        =   8
            Top             =   240
            Width           =   3495
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
         Left            =   0
         TabIndex        =   38
         ToolTipText     =   "Suministre la Dirección del Cliente."
         Top             =   1680
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
            TabIndex        =   5
            Top             =   240
            Width           =   9135
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
         Left            =   7320
         TabIndex        =   37
         ToolTipText     =   "Suministre la Cédula del Cliente."
         Top             =   840
         Width           =   2055
         Begin VB.TextBox txt_cedula 
            Alignment       =   2  'Center
            DataField       =   "cedula"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "@@.@@@.@@@"
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
            TabIndex        =   4
            Top             =   240
            Width           =   1815
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
         Left            =   2040
         TabIndex        =   36
         ToolTipText     =   "Suministre el Nº de Contrato del Cliente"
         Top             =   0
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
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   240
            Width           =   1215
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
         Left            =   0
         TabIndex        =   35
         ToolTipText     =   "Suministre Apellido y Nombre del Cliente"
         Top             =   840
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
            TabIndex        =   3
            Top             =   240
            Width           =   6975
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
         Left            =   0
         TabIndex        =   34
         ToolTipText     =   "Suministre los Teléfonos del Cliente."
         Top             =   2520
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
            TabIndex        =   6
            Top             =   240
            Width           =   2175
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
         Left            =   3720
         TabIndex        =   32
         ToolTipText     =   "Suministre la Fecha del Contrato del Cliente."
         Top             =   0
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
            TabIndex        =   33
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
            Format          =   63111169
            CurrentDate     =   39083
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
         Left            =   0
         TabIndex        =   31
         ToolTipText     =   "Suministre el Código del Cliente"
         Top             =   0
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
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame15 
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
         Left            =   5520
         TabIndex        =   30
         ToolTipText     =   "Suministre la Fecha del Contrato del Cliente."
         Top             =   3360
         Width           =   2055
         Begin MSComCtl2.DTPicker DTPicker2 
            Bindings        =   "frm_clientes.frx":090D
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
            DataSource      =   "Clientes"
            Height          =   345
            Left            =   120
            TabIndex        =   10
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
            Format          =   63111169
            CurrentDate     =   39083
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   ">>|"
         Height          =   615
         Index           =   3
         Left            =   8760
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   4440
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   ">>"
         Height          =   615
         Index           =   2
         Left            =   8040
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   4440
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "<<"
         Height          =   615
         Index           =   1
         Left            =   7320
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   4440
         Width           =   735
      End
      Begin VB.Frame Frame17 
         Height          =   1215
         Left            =   0
         TabIndex        =   27
         Top             =   4320
         Width           =   6495
         Begin VB.CommandButton cmdsalir 
            Caption         =   "C&errar"
            Height          =   660
            Left            =   5040
            TabIndex        =   16
            ToolTipText     =   "Cerrar el Sistema"
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdeliminar 
            Caption         =   "&Eliminar"
            Height          =   660
            Left            =   3840
            TabIndex        =   15
            ToolTipText     =   "Elimina de la Base de Datos a un Usuario"
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdmodificar 
            Caption         =   "&Modificar"
            Height          =   660
            Left            =   2640
            TabIndex        =   14
            ToolTipText     =   "Cambiar Característica de un Usuario"
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdguardar 
            Caption         =   "&Guardar"
            Enabled         =   0   'False
            Height          =   660
            Left            =   1440
            TabIndex        =   12
            ToolTipText     =   "Para Salvar el Usuario Agregado o Modificado"
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdcancelar 
            Caption         =   "&Cancelar"
            Height          =   660
            Left            =   240
            TabIndex        =   13
            ToolTipText     =   "Pulse este botón si desea Cancelar el Usuario Agregado"
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdagregar 
            Caption         =   "&Agregar"
            Height          =   660
            Left            =   240
            TabIndex        =   0
            TabStop         =   0   'False
            ToolTipText     =   "Botón para Agregar un Nuevo Usuario"
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "|<<"
         Height          =   615
         Index           =   0
         Left            =   6600
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   4440
         Width           =   735
      End
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2400
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "Text2"
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      DataField       =   "fecha_ini_contrato"
      DataSource      =   "Clientes"
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   360
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   7560
      TabIndex        =   18
      Top             =   960
      Width           =   7695
      Begin VB.CommandButton Busquedad_avanzada 
         Caption         =   "Búsqueda Avanzada"
         Height          =   375
         Left            =   5280
         TabIndex        =   28
         Top             =   120
         Width           =   2175
      End
      Begin MSDataListLib.DataCombo Dcmb_Buscar 
         Bindings        =   "frm_clientes.frx":0928
         Height          =   360
         Left            =   0
         TabIndex        =   19
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
   End
   Begin MSAdodcLib.Adodc status 
      Height          =   450
      Left            =   6600
      Top             =   0
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
      Height          =   450
      Left            =   4680
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
   Begin MSAdodcLib.Adodc control_clientes 
      Height          =   450
      Left            =   8760
      Top             =   0
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
      Height          =   450
      Left            =   4680
      Top             =   480
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
   Begin MSAdodcLib.Adodc clientes 
      Height          =   450
      Left            =   6840
      Top             =   480
      Visible         =   0   'False
      Width           =   2805
      _ExtentX        =   4948
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
      RecordSource    =   "tbl_clientes"
      Caption         =   "Clientes"
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
   Begin VB.Label LABEL_BUSCA 
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
      TabIndex        =   21
      Top             =   1200
      Width           =   3495
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
      Caption         =   "  CONTROL DE CLIENTES"
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
      Left            =   2760
      TabIndex        =   20
      Top             =   240
      Width           =   8655
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000000&
      BorderColor     =   &H8000000B&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   0
      Top             =   900
      Width           =   15465
   End
End
Attribute VB_Name = "frm_clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Busquedad_avanzadas_Click(Index As Integer)
            Busq_Avanzada = True
            
            clientes.CommandType = adCmdText
            
            clientes.RecordSource = "select * from tbl_clientes WHERE codigo <> '' ORDER BY codigo ASC"
            
            clientes.Refresh
            
            Call Dcmb_Buscar_Click(1)
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
    'txt_codigo.Locked = False
    'txt_contrato.Locked = False
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
    
'    cmb_cilindro.Locked = False
'    cmb_status.Locked = False
    
    fecha = DTPicker1.Value
    
    Text1.Text = DateAdd("m", 1, fec)
    
    Text2.Text = fecha
    
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
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdcancelar_Click()

On Error GoTo ControlError
    
    clientes.Recordset.CancelUpdate
'    If mvBookMark > 0 Then
'        clientes.Recordset.Bookmark = mvBookMark
'    Else
'        clientes.Recordset.MoveFirst
'    End If
'
'       clientes.Recordset.CancelUpdate
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
    Exit Sub    ' Salir para evitar el controlador.

ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "Control del Gimnasio")
        
    End Select

 End Sub

Private Sub cmdeliminar_Click()
  'esto puede producir un error si elimina el último
  'registro o el único registro del recordset

On Error GoTo ControlError
    respuesta = MsgBox("¿Desea Eliminar el Registro?", vbYesNo)
    If respuesta = vbYes Then
        clientes.Recordset.Delete
        clientes.Recordset.MoveNext
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


End Sub

Private Sub cmdguardar_Click()
Dim fec As Date
Dim ano As Date
Dim strquery As String
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
    
        MsgBox "Debe seleccionar un cilindro, por favor verifique ", vbInformation, "JerGas"
             
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
    'txtFields(2).Text = txtFields(6).Text
    'txtFields(7).Text = DateAdd("m", 1, fec)
    'Text12.Text = DateAdd("m", 1, fec)
    'ano = DTPicker2.Value
    'txtFields(8).Text = DateAdd("yyyy", 1, ano)
    'Text11.Text = DateAdd("yyyy", 1, ano)
    
    With clientes.Recordset

        mvBookMark = .Bookmark

        .Update

        .Bookmark = mvBookMark
       'Clientes.Refresh

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

Private Sub Form_Load()
'Call actualizar_cn("SQL Server")
    cmdcancelar.Visible = False
    DTPicker1.Enabled = True
    'clientes.Refresh
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
Call Mover_der(Me, Label_titulo, 0)
Call Mover_centrado(Me, Frame1)
Call Mover_der(Me, Frame3, 10)
Call Mover_der(Me, Me.LABEL_BUSCA, Frame3.Width + 15)
Shape1.Width = Me.Width
Shape1.Left = 0
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
       clientes.Recordset.MoveFirst
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
       clientes.Recordset.MovePrevious
       Command1(2).Enabled = True
       Command1(3).Enabled = True
         Me.cmdsalir.FontBold = False
         Me.cmdguardar.FontBold = False
         Me.cmdeliminar.FontBold = False
         Me.cmdcancelar.FontBold = False
         Me.cmdmodificar.FontBold = False
         Me.cmdagregar.FontBold = False
       If clientes.Recordset.BOF = True Then
        clientes.Recordset.MoveFirst
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
       clientes.Recordset.MoveNext
       Command1(0).Enabled = True
       Command1(1).Enabled = True
         Me.cmdsalir.FontBold = False
         Me.cmdguardar.FontBold = False
         Me.cmdeliminar.FontBold = False
         Me.cmdcancelar.FontBold = False
         Me.cmdmodificar.FontBold = False
         Me.cmdagregar.FontBold = False
       If clientes.Recordset.EOF = True Then
         Command1(2).Enabled = False
         Command1(3).Enabled = False
         Me.cmdsalir.FontBold = False
         Me.cmdguardar.FontBold = False
         Me.cmdeliminar.FontBold = False
         Me.cmdcancelar.FontBold = False
         Me.cmdmodificar.FontBold = False
         Me.cmdagregar.FontBold = False
         clientes.Recordset.MoveLast
       Else
       End If
    Case 3
       clientes.Recordset.MoveLast
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
            MsgBox "Error, la cédula suministrada ya existe", vbOKOnly, "Control del Cliente"
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


Private Sub habilitar_botones(valor As Boolean)
'Me.cmd_clientes.Enabled = valor
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

Private Sub txt_cedula_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
 ' End If
End Sub

Private Sub txt_direccion_KeyPress(KeyAscii As Integer)
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

Private Sub txt_observaciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Index <> 6 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 13 Or Index = 6 Or Index = 9 Then
        If ((KeyAscii < 48) Or (KeyAscii > 57)) And (KeyAscii <> 8) Then KeyAscii = 0
   End If
End Sub

