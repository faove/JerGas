VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Begin VB.Form frm_ayuda04 
   BackColor       =   &H80000004&
   Caption         =   "AYUDA AL USUARIO"
   ClientHeight    =   8490
   ClientLeft      =   4455
   ClientTop       =   1965
   ClientWidth     =   10590
   Icon            =   "frm_ayuda04.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   10590
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   7800
      Width           =   10575
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "Cerrar"
         Height          =   495
         Left            =   8880
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmd_final 
         Caption         =   ">>||"
         Height          =   495
         Left            =   7320
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmd_siguiente 
         Caption         =   ">>"
         Height          =   495
         Left            =   6360
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmd_anterior 
         Caption         =   "<<"
         Height          =   495
         Left            =   5400
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmd_inicio 
         Caption         =   "||<<"
         Height          =   495
         Left            =   4440
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "4 / 6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   170
         Width           =   735
      End
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   13150
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      FileName        =   "C:\Jergas\Ayuda\Movimientos.rtf"
      TextRTF         =   $"frm_ayuda04.frx":08CA
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   7680
      Width           =   10575
   End
End
Attribute VB_Name = "frm_ayuda04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_inicio_LostFocus()
Unload Me
End Sub

Private Sub cmd_anterior_LostFocus()
Unload Me
End Sub

Private Sub cmd_siguiente_LostFocus()
Unload Me
End Sub

Private Sub cmd_final_LostFocus()
Unload Me
End Sub

Private Sub cmd_cerrar_Click()
Unload Me
End Sub

' cuando hago click
Private Sub cmd_inicio_Click()
frm_ayuda01.Show
End Sub

Private Sub cmd_anterior_Click()
frm_ayuda03.Show
End Sub

Private Sub cmd_siguiente_Click()
frm_ayuda05.Show
End Sub

Private Sub cmd_final_Click()
frm_ayuda06.Show
End Sub










Private Sub cmd_inicio_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_inicio.FontBold = True
Me.cmd_anterior.FontBold = False
Me.cmd_siguiente.FontBold = False
Me.cmd_final.FontBold = False
Me.cmd_cerrar.FontBold = False
End Sub

Private Sub cmd_anterior_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_inicio.FontBold = False
Me.cmd_anterior.FontBold = True
Me.cmd_siguiente.FontBold = False
Me.cmd_final.FontBold = False
Me.cmd_cerrar.FontBold = False
End Sub

Private Sub cmd_siguiente_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_inicio.FontBold = False
Me.cmd_anterior.FontBold = False
Me.cmd_siguiente.FontBold = True
Me.cmd_final.FontBold = False
Me.cmd_cerrar.FontBold = False
End Sub

Private Sub cmd_final_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_inicio.FontBold = False
Me.cmd_anterior.FontBold = False
Me.cmd_siguiente.FontBold = False
Me.cmd_final.FontBold = True
Me.cmd_cerrar.FontBold = False
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_inicio.FontBold = False
Me.cmd_anterior.FontBold = False
Me.cmd_siguiente.FontBold = False
Me.cmd_final.FontBold = False
Me.cmd_cerrar.FontBold = True
End Sub



