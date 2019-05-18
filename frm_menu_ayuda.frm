VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm_menu_ayuda 
   BackColor       =   &H00C0FFFF&
   Caption         =   "SISTEMA DE GESTIÓN Y CONTROL JER-GAS, C.A."
   ClientHeight    =   8490
   ClientLeft      =   10365
   ClientTop       =   1965
   ClientWidth     =   4695
   ControlBox      =   0   'False
   Icon            =   "frm_menu_ayuda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   4695
   Begin VB.OptionButton Option7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ACERCA DE"
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
      TabIndex        =   8
      Top             =   3000
      Width           =   2775
   End
   Begin VB.OptionButton Option6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "HERRAMIENTAS"
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
      TabIndex        =   7
      Top             =   2640
      Width           =   2775
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "REPORTES"
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
      TabIndex        =   6
      Top             =   2280
      Width           =   2775
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "MOVIMIENTOS"
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
      TabIndex        =   5
      Top             =   1920
      Width           =   2775
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ARCHIVOS"
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
      TabIndex        =   4
      Top             =   1560
      Width           =   2775
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "MENÚ PRINCIPAL"
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
      TabIndex        =   3
      Top             =   1200
      Width           =   2775
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "INTRODUCCIÓN"
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
      TabIndex        =   2
      Top             =   840
      Width           =   2775
   End
   Begin MSForms.Frame Frame1 
      Height          =   3495
      Left            =   240
      OleObjectBlob   =   "frm_menu_ayuda.frx":08CA
      TabIndex        =   1
      Top             =   240
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
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
      Left            =   1320
      TabIndex        =   0
      Top             =   7680
      Width           =   1695
   End
End
Attribute VB_Name = "frm_menu_ayuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Option1_Click()
frm_ayuda01.Show
End Sub

Private Sub Option2_Click()
frm_ayuda02.Show
End Sub

Private Sub Option3_Click()
frm_ayuda03.Show
End Sub

Private Sub Option4_Click()
frm_ayuda04.Show
End Sub

Private Sub Option5_Click()
frm_ayuda05.Show
End Sub

Private Sub Option6_Click()
frm_ayuda06.Show
End Sub

