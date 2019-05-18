VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frm_respaldo_mensual 
   Caption         =   "SISTEMA DE GESTIÓN Y CONTROL JER-GAS, C.A."
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10020
   Icon            =   "frm_respaldo_mensual.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   10020
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdcerrar 
      Caption         =   "SALIR"
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
      Left            =   4440
      TabIndex        =   24
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PROCESAR"
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
      Left            =   4440
      TabIndex        =   23
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Archivo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   3840
      TabIndex        =   18
      Top             =   1920
      Width           =   2400
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Resumen de Procesos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1935
      Left            =   3840
      TabIndex        =   14
      Top             =   2820
      Visible         =   0   'False
      Width           =   2415
      Begin VB.TextBox TotalArchivos 
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
         Height          =   285
         Left            =   1800
         TabIndex        =   21
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox eliminados 
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
         Height          =   285
         Left            =   1800
         TabIndex        =   20
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox copiados 
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
         Height          =   285
         Left            =   1800
         TabIndex        =   19
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total de Archivos"
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
         Height          =   195
         Left            =   100
         TabIndex        =   17
         Top             =   1500
         Width           =   1515
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "No Encontrados"
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
         Height          =   195
         Left            =   100
         TabIndex        =   16
         Top             =   910
         Width           =   1380
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Transferidos"
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
         Height          =   195
         Left            =   100
         TabIndex        =   15
         Top             =   360
         Width           =   1065
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ARCHIVOS DE ORIGEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   4575
      Left            =   360
      TabIndex        =   7
      Top             =   1800
      Width           =   3375
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Disco"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   3105
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   105
            TabIndex        =   13
            Top             =   240
            Width           =   2940
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Carpeta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1530
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   3105
         Begin VB.DirListBox Dir1 
            Height          =   1215
            Left            =   90
            TabIndex        =   11
            Top             =   255
            Width           =   2970
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Archivos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1860
         Left            =   120
         TabIndex        =   8
         Top             =   2640
         Width           =   3120
         Begin VB.FileListBox File1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1560
            Left            =   45
            TabIndex        =   9
            Top             =   225
            Width           =   3000
         End
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ARCHIVOS DE DESTINOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   4575
      Left            =   6360
      TabIndex        =   0
      Top             =   1800
      Width           =   3375
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Disco"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3105
         Begin VB.DriveListBox Drive2 
            Height          =   315
            Left            =   105
            TabIndex        =   6
            Top             =   225
            Width           =   2940
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Carpeta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1530
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   3105
         Begin VB.DirListBox Dir2 
            Height          =   1215
            Left            =   105
            TabIndex        =   4
            Top             =   255
            Width           =   2970
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Archivos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1860
         Left            =   120
         TabIndex        =   1
         Top             =   2640
         Width           =   3120
         Begin VB.FileListBox File2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1560
            Left            =   45
            TabIndex        =   2
            Top             =   210
            Width           =   3000
         End
      End
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   4080
      TabIndex        =   22
      Top             =   2400
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   0
      Max             =   600
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   8760
      Picture         =   "frm_respaldo_mensual.frx":08CA
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   8160
      Picture         =   "frm_respaldo_mensual.frx":1194
      Top             =   720
      Width           =   480
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000001&
      BorderColor     =   &H00C00000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   0
      Top             =   780
      Width           =   8070
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   9360
      Picture         =   "frm_respaldo_mensual.frx":1A5E
      Top             =   720
      Width           =   480
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000001&
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   0
      Top             =   720
      Width           =   8055
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   0
      Top             =   720
      Width           =   10695
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   5655
      Left            =   0
      Top             =   1080
      Width           =   10095
   End
End
Attribute VB_Name = "frm_respaldo_mensual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcerrar_Click()
Unload Me
End Sub

Private Sub Command2_Click()

    Dim I, control As Integer
    Dim Origen As String
    Dim Destino As String
    
    ProgressBar1.Visible = True
        
    TotalArchivos.Text = File1.ListCount
    Archivo.Text = ""
    eliminados.Text = ""
    copiados.Text = ""
    
   On Error Resume Next
   On Error GoTo ErrorVacio
ErrorVacio:        If (Err.Number = 53) Then
                            Dim Aux As String
                            Aux = File1.List(I)
                            MsgBox "Archivo No Encontrado " & Aux, vbExclamation, "Control de Archivos"
                            control = control + 1
                            eliminados.Text = control
                            Dir1_Change
                        End If
                        If (Err.Number = 380) Or (Err.Number = 76) Then
                            MsgBox "Carpeta No Encontrado ", vbExclamation, "Control de Carpetas"
'                            Exit Sub
                        End If
                        If (Err.Number = 75) Then
                            MsgBox "Archivo Existente en el Destino : " & File1.List(I), vbExclamation, "Control de Archivos"
                            Exit Sub
                        End If
                        If (Err.Number = 61) Then
                            MsgBox "Disco de Destino LLeno: ", vbExclamation, "Control de Archivos"
                            Exit Sub
                        End If
                        If (Err.Number = 52) Then
                            MsgBox "Error en tipo de Archivo: ", vbExclamation, "Control de Archivos"
                            Exit Sub
                        End If
    
    ProgressBar1.Value = 0
    ProgressBar1.Min = 0
    Origen = File1.Path
    If Right(Origen, 1) <> "\" Then Origen = Origen & "\"
    Destino = Dir2.Path
    If Right(Destino, 1) <> "\" Then Destino = Destino & "\"

    For I = 0 To File1.ListCount - 1
        DoEvents
        FileCopy Origen & File1.List(I), Destino & File1.List(I)
        ProgressBar1.Value = I
        Archivo.Text = File1.List(I)
        File2.Refresh
    Next I
    Me.MousePointer = 0
    MsgBox "Transferido : " & File1.ListCount & "  Archivos", vbInformation, "Transferencia Completada"
    copiados.Text = File1.ListCount
    Frame9.Visible = True
    Me.Archivo.Text = ""
    Me.eliminados.Text = 0
End Sub
Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    File1.Refresh
End Sub
Private Sub Dir2_Change()
    File2.Path = Dir2.Path
    File2.Refresh
End Sub
Private Sub Drive1_Change()
     
            On Error Resume Next
            On Error GoTo ErrorVacio
    
ErrorVacio:          If (Err.Number = 68) Or (Err.Number = 70) Then
                            MsgBox "Dispositivo NO disponible: ", vbExclamation, "Control de Dispositivos"
                            Exit Sub
                        End If
                        
    Dir1.Path = Drive1.Drive
    Dir1.Refresh
    Image1.ToolTipText = Drive1.Drive
End Sub
Private Sub Drive2_Change()
  
            On Error Resume Next
            On Error GoTo ErrorVacio
    
ErrorVacio:       If (Err.Number = 68) Or (Err.Number = 70) Then
                            MsgBox "Dispositivo NO disponible: ", vbExclamation, "Control de Dispositivos"
                            Exit Sub
                        End If
                        
    Dir2.Path = Drive2.Drive
    
End Sub

Private Sub Timer1_Timer()
Drive1_Change
Drive2_Change
End Sub


