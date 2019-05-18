VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Begin VB.Form frm_ayuda03 
   BackColor       =   &H80000004&
   Caption         =   "AYUDA AL USUARIO"
   ClientHeight    =   8490
   ClientLeft      =   4455
   ClientTop       =   1965
   ClientWidth     =   10590
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
      Left            =   -120
      TabIndex        =   1
      Top             =   7800
      Width           =   10575
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "Cerrar"
         Height          =   495
         Left            =   8880
         TabIndex        =   6
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   ">>|"
         Height          =   495
         Index           =   3
         Left            =   7560
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   ">>"
         Height          =   495
         Index           =   2
         Left            =   6720
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "<<"
         Height          =   495
         Index           =   1
         Left            =   5880
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "|<<"
         Height          =   495
         Index           =   0
         Left            =   5040
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   120
         Width           =   735
      End
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
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frm_ayuda03.frx":0000
   End
End
Attribute VB_Name = "frm_ayuda03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
