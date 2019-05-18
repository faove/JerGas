VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm_productos 
   Caption         =   "Productos"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   7110
   Begin VB.Frame Frame4 
      Caption         =   "Stock Actual"
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
      Height          =   615
      Left            =   4440
      TabIndex        =   18
      ToolTipText     =   "Suministre la Cédula o el Código del Usuario."
      Top             =   1200
      Width           =   1695
      Begin MSForms.TextBox TextBox3 
         Bindings        =   "frm_productos.frx":0000
         DataField       =   "contrato_num"
         DataSource      =   "Ado_Clientes"
         Height          =   300
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1455
         VariousPropertyBits=   746604571
         Size            =   "2566;529"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   240
      TabIndex        =   10
      Top             =   3360
      Width           =   6615
      Begin VB.CommandButton cmdmodificar 
         Caption         =   "&Modificar"
         Height          =   660
         Left            =   3360
         TabIndex        =   17
         ToolTipText     =   "Cambiar Característica de un Usuario"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdbuscar 
         Caption         =   "&Buscar"
         Height          =   660
         Left            =   2280
         TabIndex        =   16
         ToolTipText     =   "Busca un Usuario"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "&Cancelar"
         Height          =   660
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Pulse este botón si desea Cancelar el Usuario Agregado"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Salir"
         Height          =   660
         Left            =   5520
         TabIndex        =   14
         ToolTipText     =   "Cerrar el Sistema"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Guardar"
         Height          =   660
         Left            =   1200
         TabIndex        =   13
         ToolTipText     =   "Para Salvar el Usuario Agregado o Modificado"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Eliminar"
         Height          =   660
         Left            =   4440
         TabIndex        =   12
         ToolTipText     =   "Elimina de la Base de Datos a un Usuario"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Agregar"
         Height          =   660
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Botón para Agregar un Nuevo Usuario"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Código"
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
      Height          =   615
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "Suministre la Cédula o el Código del Usuario."
      Top             =   360
      Width           =   1095
      Begin MSForms.TextBox TextBox1 
         Bindings        =   "frm_productos.frx":002D
         DataField       =   "codigo"
         DataSource      =   "Ado_Clientes"
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   855
         VariousPropertyBits=   746604571
         Size            =   "1508;529"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Fecha de Ingreso"
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
      Height          =   615
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "Suministre la Fecha de Inscripción del Año que se está Cancelando."
      Top             =   1200
      Width           =   2055
      Begin MSComCtl2.DTPicker DTPicker2 
         Bindings        =   "frm_productos.frx":0054
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   3
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         Format          =   70057985
         CurrentDate     =   37240
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Descripción"
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
      Height          =   615
      Left            =   1560
      TabIndex        =   4
      ToolTipText     =   "Suministre el Nombre y Apellido del Usuario."
      Top             =   360
      Width           =   3615
      Begin MSForms.TextBox TextBox2 
         Bindings        =   "frm_productos.frx":0087
         DataField       =   "cliente"
         DataSource      =   "Ado_Clientes"
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3375
         VariousPropertyBits=   746604571
         Size            =   "5953;529"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Stock Inicial"
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
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      ToolTipText     =   "Suministre la Cédula o el Código del Usuario."
      Top             =   1200
      Width           =   1695
      Begin MSForms.TextBox TextBox5 
         Bindings        =   "frm_productos.frx":00AF
         DataField       =   "contrato_num"
         DataSource      =   "Ado_Clientes"
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1455
         VariousPropertyBits=   746604571
         Size            =   "2566;529"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame9 
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
      Height          =   615
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Suministre el Nombre y Apellido del Usuario."
      Top             =   2040
      Width           =   4815
      Begin MSForms.TextBox TextBox7 
         Bindings        =   "frm_productos.frx":00DC
         DataField       =   "direccion"
         DataSource      =   "Ado_Clientes"
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4575
         VariousPropertyBits=   746604571
         Size            =   "8070;529"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "frm_productos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
