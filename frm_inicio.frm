VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frm_inicio 
   BackColor       =   &H80000004&
   Caption         =   "Inicio de Sesión"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8535
   Icon            =   "frm_inicio.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_cancelar 
      BackColor       =   &H8000000A&
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmd_aceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   18
      Top             =   5040
      Width           =   8535
   End
   Begin VB.TextBox TXT_CONTROL 
      DataField       =   "status"
      DataSource      =   "TBL_CONTROL"
      Height          =   285
      Left            =   1320
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   7440
      Width           =   975
   End
   Begin VB.TextBox TXT_ID 
      DataField       =   "Id"
      DataSource      =   "TBL_CONTROL"
      Height          =   285
      Left            =   2400
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   7440
      Width           =   975
   End
   Begin VB.Frame Grupo_Usuario 
      BackColor       =   &H80000004&
      Height          =   1815
      Left            =   1560
      TabIndex        =   4
      Top             =   3000
      Width           =   5415
      Begin VB.TextBox txt_flag 
         DataField       =   "status"
         DataSource      =   "tbl_usuario"
         Height          =   285
         Left            =   4080
         TabIndex        =   8
         Top             =   2040
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txt_id_usuario 
         DataField       =   "id_usuario"
         DataSource      =   "tbl_usuario"
         Height          =   285
         Left            =   3840
         TabIndex        =   7
         Top             =   1680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txt_pass 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   480
         PasswordChar    =   " "
         TabIndex        =   1
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txt_path 
         DataField       =   "path_foto"
         DataSource      =   "tbl_usuario"
         Height          =   285
         Left            =   3240
         TabIndex        =   6
         Top             =   2040
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   3480
         Top             =   1680
      End
      Begin VB.TextBox txt_grupo_usuario 
         DataField       =   "id_grupo"
         DataSource      =   "tbl_usuario"
         Height          =   285
         Left            =   4920
         TabIndex        =   5
         Top             =   1680
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSAdodcLib.Adodc tbl_usuario 
         Height          =   375
         Left            =   240
         Top             =   2160
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
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
         RecordSource    =   "select * from tbl_usuario where status = '1' order by nombre_usuario"
         Caption         =   "tbl_usuario"
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
      Begin MSDataListLib.DataCombo DCombo_login 
         Bindings        =   "frm_inicio.frx":08CA
         Height          =   315
         Left            =   480
         TabIndex        =   0
         Top             =   480
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "nombre_usuario"
         BoundColumn     =   "nombre_usuario"
         Text            =   ""
      End
      Begin MSForms.Label Label1 
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   840
         Width           =   1095
         ForeColor       =   -2147483635
         BackColor       =   16777215
         Caption         =   "Password:"
         Size            =   "1931;450"
         BorderColor     =   -2147483635
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label login 
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   240
         Width           =   1095
         ForeColor       =   -2147483635
         BackColor       =   16777215
         Caption         =   "Usuario:"
         Size            =   "1931;450"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Image imgProf 
         BorderStyle     =   1  'Fixed Single
         Height          =   1215
         Left            =   3840
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1335
      End
      Begin MSForms.Label Label4 
         Height          =   255
         Left            =   3840
         TabIndex        =   9
         Top             =   240
         Width           =   1335
         ForeColor       =   -2147483635
         BackColor       =   16777215
         Caption         =   "Foto:"
         Size            =   "2355;450"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin MSAdodcLib.Adodc TBL_CONTROL 
      Height          =   375
      Left            =   6360
      Top             =   0
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      RecordSource    =   "tbl_status"
      Caption         =   "TBLCONTROL"
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
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   975
      Left            =   2640
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   480
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1720
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      Appearance      =   0
      FileName        =   "C:\Jergas\Imagen\jergas.rtf"
      TextRTF         =   $"frm_inicio.frx":08E4
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   4920
      Width           =   8535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Caption         =   "SISTEMA DE "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   1680
      TabIndex        =   17
      Top             =   1560
      Width           =   5175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Caption         =   "CONTROL Y GESTIÓN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   1680
      TabIndex        =   16
      Top             =   2040
      Width           =   5175
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000004&
      Caption         =   "Versión 1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   15
      Top             =   2520
      Width           =   1695
   End
End
Attribute VB_Name = "frm_inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_Click()
On Error GoTo ControlError

Dim strquery

If Me.DCombo_login.Text = "" Or IsNull(Me.DCombo_login.Text) Then
    
    MsgBox "Por favor, suministre el nombre del usuario, gracias", vbCritical, "JerGas C.A."
    DCombo_login.SetFocus
    Exit Sub
    
End If


With tbl_usuario
        
    .CommandType = adCmdText
    
    strquery = "select * from tbl_usuario where nombre_usuario = '" & Me.DCombo_login.Text & "' and clave_usuario = '" & Me.txt_pass.Text & "' and status = '1'"
    
    .RecordSource = strquery
            
    .Refresh

    
    If .Recordset.EOF Then
    
        MsgBox "Verifique el password suministrado", vbOKOnly, "JerGas C.A."
        
        strquery = "select * from tbl_usuario where nombre_usuario = '" & Me.DCombo_login.Text & "' and status = '1' ORDER BY nombre_usuario"
    
        .RecordSource = strquery
            
        .Refresh
        
        Me.txt_pass.Text = ""
        
        Me.DCombo_login.SetFocus
        
    Else
        If Me.txt_flag.Text = -1 Then
        
            Me.txt_pass.Enabled = False
            Me.DCombo_login.Enabled = False
            Me.cmd_aceptar.Enabled = False
            Me.txt_pass.SetFocus
            
            Timer2.Interval = 80
            
            Exit Sub
        
        End If
        '----------------------------------------------------------
        'Variable utilizada para identificar al usuario del sistema
        '----------------------------------------------------------
        Usuario = Me.txt_id_usuario.Text
        user_name = Me.DCombo_login.Text
        user_grupo = Me.txt_grupo_usuario.Text
                
        Unload Me
        
        '-----------------------------------------
        'Llamada a la pantalla principal de SIAGEP
        '-----------------------------------------
        MDIForm1.Show
        
    End If
    
End With
    
    Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            MsgBox "Formato No Válido", vbOKOnly, "JerGas C.A."
        Case 3001
            MsgBox "Verifique el password suministrado", vbOKOnly, "JerGas C.A."
    End Select
End Sub

Private Sub cmd_cancelar_Click()
End
End Sub



Private Sub Timer1_Timer()
On Error GoTo control_error

'    If Format(Date, "dd/mm/yyyy") >= "30/11/2008" Then
        
'            TBL_CONTROL.CommandType = adCmdText
            
'            TBL_CONTROL.RecordSource = "select * from tbl_status WHERE id = 5"
            
'            TBL_CONTROL.Refresh
            
'            If TBL_CONTROL.Recordset.EOF Then
'                MsgBox "Error en el sistema de Control", vbCritical
'                End
'            End If
'            If TBL_CONTROL.Recordset!status = "VI" Then
'                TBL_CONTROL.Recordset!status = "Vl"
'                TBL_CONTROL.Recordset.Update
'            Else
'                End
'            End If
            
'    End If
'        If Format(Date, "dd/mm/yyyy") < "30/11/2007" Then
        
'            TBL_CONTROL.CommandType = adCmdText
            
'            TBL_CONTROL.RecordSource = "select * from tbl_status WHERE id = 5"
            
'            TBL_CONTROL.Refresh
            
'            If TBL_CONTROL.Recordset.EOF Then
'                MsgBox "Error en el sistema de Control", vbCritical
'                End
'            End If
'            If TBL_CONTROL.Recordset!status = "Vl" Then
'                End
'            End If
            
'    End If

    imgProf.Picture = LoadPicture(Me.txt_path.Text)


Exit Sub
control_error:
        Select Case Err.Number
            Case 13
                MsgBox ("Error en los datos 10")

        End Select
    Exit Sub
End Sub
