VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm_conf_despacho 
   Caption         =   "Asignaci�n de Despacho"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9825
   Icon            =   "frm_conf_despacho.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11.4
   ScaleMode       =   0  'User
   ScaleWidth      =   12
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   5055
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   9495
      Begin VB.CommandButton Command1 
         Caption         =   ">>|"
         Height          =   615
         Index           =   3
         Left            =   3000
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   4080
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   ">>"
         Height          =   615
         Index           =   2
         Left            =   2280
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   4080
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "<<"
         Height          =   615
         Index           =   1
         Left            =   1560
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   4080
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "|<<"
         Height          =   615
         Index           =   0
         Left            =   840
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   4080
         Width           =   735
      End
      Begin VB.Frame Frame33 
         Caption         =   "CONFIGURACI�N DE DESPACHO"
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
         Left            =   4800
         TabIndex        =   19
         Top             =   120
         Width           =   4455
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            DataField       =   "Nro_pedidos_x_chofer"
            DataSource      =   "control_procesos"
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
            Left            =   3360
            TabIndex        =   21
            Text            =   "30"
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "N�meros de Pedidos M�ximos por Cami�n:"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   480
            Width           =   3255
         End
      End
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "Cerrar"
         Height          =   735
         Left            =   7320
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   4080
         Width           =   1455
      End
      Begin VB.CommandButton cmd_guardar 
         Caption         =   "Guardar"
         Height          =   735
         Left            =   5640
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Caption         =   "DATOS DEL CHOFER"
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
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   4455
         Begin MSDataListLib.DataCombo Dcmb_ruta 
            DataField       =   "id_ruta"
            DataSource      =   "choferes"
            Height          =   315
            Left            =   1680
            TabIndex        =   17
            Top             =   600
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSForms.TextBox TextBox11 
            Bindings        =   "frm_conf_despacho.frx":08CA
            DataField       =   "telefono_cel"
            DataSource      =   "choferes"
            Height          =   300
            Left            =   2280
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   3240
            Width           =   1935
            VariousPropertyBits=   746604569
            Size            =   "3413;529"
            FontEffects     =   1073750016
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label11 
            Caption         =   "Tel�fono Celular"
            Height          =   240
            Left            =   2280
            TabIndex        =   15
            Top             =   3000
            Width           =   1455
         End
         Begin MSForms.TextBox TextBox10 
            Bindings        =   "frm_conf_despacho.frx":08F7
            DataField       =   "telefono_hab"
            DataSource      =   "choferes"
            Height          =   300
            Left            =   120
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   3240
            Width           =   1935
            VariousPropertyBits=   746604569
            Size            =   "3413;529"
            FontEffects     =   1073750016
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label10 
            Caption         =   "Tel�fono Habitaci�n"
            Height          =   240
            Left            =   120
            TabIndex        =   13
            Top             =   3000
            Width           =   1575
         End
         Begin MSForms.TextBox TextBox4 
            Bindings        =   "frm_conf_despacho.frx":0924
            DataField       =   "cedula"
            DataSource      =   "choferes"
            Height          =   300
            Left            =   120
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1800
            Width           =   2175
            VariousPropertyBits=   746604569
            Size            =   "3836;529"
            FontEffects     =   1073750016
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label5 
            Caption         =   "C�dula de Identidad"
            Height          =   240
            Left            =   120
            TabIndex        =   11
            Top             =   1560
            Width           =   1575
         End
         Begin MSForms.TextBox txt_direccion 
            Bindings        =   "frm_conf_despacho.frx":0951
            DataField       =   "direccion"
            DataSource      =   "choferes"
            Height          =   540
            Left            =   120
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   2400
            Width           =   4095
            VariousPropertyBits=   -1400879079
            Size            =   "7223;952"
            FontEffects     =   1073750016
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label4 
            Caption         =   "Direcci�n"
            Height          =   240
            Left            =   120
            TabIndex        =   9
            Top             =   2160
            Width           =   1575
         End
         Begin MSForms.TextBox txt_nombre 
            Bindings        =   "frm_conf_despacho.frx":097E
            DataField       =   "nombre"
            DataSource      =   "choferes"
            Height          =   300
            Left            =   120
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1200
            Width           =   3375
            VariousPropertyBits=   746604569
            Size            =   "5953;529"
            FontEffects     =   1073750016
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label3 
            Caption         =   "Apellidos y Nombres"
            Height          =   240
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   1575
         End
         Begin MSForms.TextBox txt_id_chofer 
            Bindings        =   "frm_conf_despacho.frx":09AB
            DataField       =   "id_chofer"
            DataSource      =   "choferes"
            Height          =   300
            Left            =   120
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   600
            Width           =   1335
            VariousPropertyBits=   746604569
            Size            =   "2355;529"
            FontEffects     =   1073750016
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            Caption         =   "C�digo del Chofer"
            Height          =   240
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "N� de la Ruta"
            Height          =   240
            Left            =   1680
            TabIndex        =   4
            Top             =   360
            Width           =   1095
         End
      End
      Begin MSAdodcLib.Adodc choferes 
         Height          =   495
         Left            =   3960
         Top             =   4680
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
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
         Caption         =   "choferes"
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
   End
   Begin MSAdodcLib.Adodc ruta 
      Height          =   375
      Left            =   3000
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
      RecordSource    =   "tbl_ruta"
      Caption         =   "ruta"
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
   Begin MSAdodcLib.Adodc control_procesos 
      Height          =   375
      Left            =   5160
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
      RecordSource    =   "tbl_control_procesos"
      Caption         =   "control_procesos"
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
   Begin VB.Label Label_titulo 
      BackColor       =   &H80000001&
      Caption         =   "  CONTROL DE DESPACHOS"
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
      Left            =   2160
      TabIndex        =   0
      Top             =   600
      Width           =   7695
   End
End
Attribute VB_Name = "frm_conf_despacho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cerrar_Click()
Unload Me
End Sub

Private Sub txt_precio_cilind_Change()

End Sub

Private Sub Command1_Click(Index As Integer)
On Error GoTo ControlError

Select Case Index
    Case 0
       choferes.Recordset.MoveFirst
       Command1(0).Enabled = False
       Command1(1).Enabled = False
       Command1(2).Enabled = True
       Command1(3).Enabled = True
         Me.cmd_guardar.FontBold = False
         Me.cmd_cerrar.FontBold = False
    Case 1
       choferes.Recordset.MovePrevious
       Command1(2).Enabled = True
       Command1(3).Enabled = True
         Me.cmd_guardar.FontBold = False
         Me.cmd_cerrar.FontBold = False
       If choferes.Recordset.BOF = True Then
        choferes.Recordset.MoveFirst
        Command1(0).Enabled = False
        Command1(1).Enabled = False
         Me.cmd_guardar.FontBold = False
         Me.cmd_cerrar.FontBold = False
       Else
        End If
    Case 2
       choferes.Recordset.MoveNext
       Command1(0).Enabled = True
       Command1(1).Enabled = True
         Me.cmd_guardar.FontBold = False
         Me.cmd_cerrar.FontBold = False
       If choferes.Recordset.EOF = True Then
         Command1(2).Enabled = False
         Command1(3).Enabled = False
         Me.cmd_guardar.FontBold = False
         Me.cmd_cerrar.FontBold = False
         choferes.Recordset.MoveLast
       Else
       End If
    Case 3
       choferes.Recordset.MoveLast
       Command1(0).Enabled = True
       Command1(1).Enabled = True
       Command1(2).Enabled = False
       Command1(3).Enabled = False
         Me.cmd_guardar.FontBold = False
         Me.cmd_cerrar.FontBold = False
End Select

ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Eval�a el n�mero de error.
'        Case 2021
'            v = MsgBox("Formato No V�lido", vbOKOnly, "Control del Cliente")
'        Case 3314
'            MsgBox "Verifique la C�dula ", vbOKOnly, "Control del Cliente"
'        Case 524
'            MsgBox "Verifique el Nombre y Apellido  ", vbOKOnly, "Control del Cliente"
        Case -2147467259
            MsgBox "Error, la c�dula suministrada ya existe", vbOKOnly, "Control de Clientes"
'        Case -2147217842
'            MsgBox "Error, cancele la operaci�n y vuelva a intentarlo", vbOKOnly, "Control del Cliente"
'        Case -2147217887
'            MsgBox "Error, al crear hist�rico, se recomienda borrar el registro y volverlo a crear", vbOKOnly, "Control del Cliente"
    End Select
End Sub
