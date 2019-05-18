VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_resumen_pedidos_estantes 
   Caption         =   "SISTEMA DE GESTIÓN Y CONTROL "
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7590
   Icon            =   "frm_resumen_pedidos_estantes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "FECHA INICIAL"
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
      Left            =   720
      TabIndex        =   4
      Top             =   1320
      Width           =   2175
      Begin MSComCtl2.DTPicker DTPicker_desde 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   65339393
         CurrentDate     =   39344
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CERRAR"
      Height          =   735
      Left            =   5520
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROCESAR"
      Height          =   735
      Left            =   5520
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "FECHA FINAL"
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
      Left            =   3000
      TabIndex        =   0
      Top             =   1320
      Width           =   2175
      Begin MSComCtl2.DTPicker DTPicker_hasta 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   65339393
         CurrentDate     =   39365
      End
   End
   Begin VB.Label Label_titulo 
      BackColor       =   &H80000001&
      Caption         =   "  RESUMEN MENSUAL DE VENTAS (ESTANTES)"
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
      Left            =   720
      TabIndex        =   6
      Top             =   480
      Width           =   8655
   End
End
Attribute VB_Name = "frm_resumen_pedidos_estantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
rpt_resumen_pedidos_estantes.Show
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

