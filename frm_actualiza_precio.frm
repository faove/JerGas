VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_actualiza_precio 
   BackColor       =   &H80000013&
   Caption         =   " SISTEMA DE GESTIÓN Y CONTROL JER-GAS, C.A."
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8295
   Icon            =   "frm_actualiza_precio.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
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
      Height          =   1125
      Left            =   6360
      TabIndex        =   21
      ToolTipText     =   "Suministre la Cédula o el Código del Usuario."
      Top             =   1080
      Width           =   1455
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         DataField       =   "Alicuota"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   " #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "instalacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000013&
         Caption         =   "Alicuota I.V.A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox Text3 
      DataField       =   "id_inst"
      DataSource      =   "materiales"
      Height          =   285
      Left            =   5640
      TabIndex        =   15
      Text            =   "Text3"
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSAdodcLib.Adodc materiales 
      Height          =   375
      Left            =   2040
      Top             =   120
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
      RecordSource    =   "tbl_materiales"
      Caption         =   "materiales"
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
   Begin VB.CommandButton Command1 
      Caption         =   "|<<"
      Height          =   495
      Index           =   0
      Left            =   2520
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<<"
      Height          =   495
      Index           =   1
      Left            =   3360
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      Height          =   495
      Index           =   2
      Left            =   4200
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>|"
      Height          =   495
      Index           =   3
      Left            =   5040
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2400
      Width           =   735
   End
   Begin MSAdodcLib.Adodc instalacion 
      Height          =   375
      Left            =   120
      Top             =   120
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
      RecordSource    =   "tbl_instalacion"
      Caption         =   "instalacion"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
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
      Height          =   1125
      Left            =   360
      TabIndex        =   6
      ToolTipText     =   "Suministre la Cédula o el Código del Usuario."
      Top             =   1080
      Width           =   5895
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         DataField       =   "precio_instalacion_2"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   " #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "instalacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   18
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         DataField       =   "precio_instalacion_1"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   " #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "instalacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   16
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         DataField       =   "precio_cilindro"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   " #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "instalacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         DataField       =   "id_inst"
         DataSource      =   "instalacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   500
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Caption         =   "Nuevas Instalaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   20
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Caption         =   "2 Cil"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   19
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Caption         =   "1 Cil"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   17
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000013&
         Caption         =   "Precio de Venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         Caption         =   "Cilindro (Kgs)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame17 
      BackColor       =   &H80000013&
      ForeColor       =   &H80000001&
      Height          =   975
      Left            =   1440
      TabIndex        =   5
      Top             =   2880
      Width           =   5535
      Begin VB.CommandButton cmdcancelar 
         Caption         =   "&Cancelar"
         Height          =   660
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Pulse este botón si desea Cancelar el Usuario Agregado"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdguardar 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   660
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Para Salvar el Usuario Agregado o Modificado"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdmodificar 
         Caption         =   "&Modificar"
         Height          =   660
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Cambiar Característica de un Usuario"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdsalir 
         Caption         =   "C&errar"
         Height          =   660
         Left            =   4080
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cerrar el Sistema"
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label Label_titulo 
      BackColor       =   &H80000001&
      Caption         =   "  ACTUALIZAR  PRECIOS DE VENTA"
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
      Left            =   2280
      TabIndex        =   7
      Top             =   480
      Width           =   6135
   End
End
Attribute VB_Name = "frm_actualiza_precio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
On Error GoTo ControlError

Select Case Index
    Case 0
       instalacion.Recordset.MoveFirst
       Command1(0).Enabled = False
       Command1(1).Enabled = False
       Command1(2).Enabled = True
       Command1(3).Enabled = True
         Me.cmdsalir.FontBold = False
         Me.cmdguardar.FontBold = False
         Me.cmdcancelar.FontBold = False
         Me.cmdmodificar.FontBold = False
    Case 1
       instalacion.Recordset.MovePrevious
       Command1(2).Enabled = True
       Command1(3).Enabled = True
         Me.cmdsalir.FontBold = False
         Me.cmdguardar.FontBold = False
         Me.cmdcancelar.FontBold = False
         Me.cmdmodificar.FontBold = False
       
       If instalacion.Recordset.BOF = True Then
        instalacion.Recordset.MoveFirst
        Command1(0).Enabled = False
        Command1(1).Enabled = False
         Me.cmdsalir.FontBold = False
         Me.cmdguardar.FontBold = False
         Me.cmdcancelar.FontBold = False
         Me.cmdmodificar.FontBold = False
      Else
        End If
    Case 2
       instalacion.Recordset.MoveNext
       Command1(0).Enabled = True
       Command1(1).Enabled = True
         Me.cmdsalir.FontBold = False
         Me.cmdguardar.FontBold = False
         Me.cmdcancelar.FontBold = False
         Me.cmdmodificar.FontBold = False
       
       If instalacion.Recordset.EOF = True Then
         Command1(2).Enabled = False
         Command1(3).Enabled = False
         Me.cmdsalir.FontBold = False
         Me.cmdguardar.FontBold = False
         Me.cmdcancelar.FontBold = False
         Me.cmdmodificar.FontBold = False

         instalacion.Recordset.MoveLast
       Else
       End If
    Case 3
       instalacion.Recordset.MoveLast
       Command1(0).Enabled = True
       Command1(1).Enabled = True
       Command1(2).Enabled = False
       Command1(3).Enabled = False
         Me.cmdsalir.FontBold = False
         Me.cmdguardar.FontBold = False
         Me.cmdcancelar.FontBold = False
         Me.cmdmodificar.FontBold = False

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
            MsgBox "Error, el Código suministrado ya existe", vbOKOnly, "Control del Cliente"
'        Case -2147217842
'            MsgBox "Error, cancele la operación y vuelva a intentarlo", vbOKOnly, "Control del Cliente"
'        Case -2147217887
'            MsgBox "Error, al crear histórico, se recomienda borrar el registro y volverlo a crear", vbOKOnly, "Control del Cliente"
    End Select
End Sub

Private Sub cmdmodificar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdmodificar.FontBold = True
Me.cmdguardar.FontBold = False
Me.cmdcancelar.FontBold = False
Me.cmdsalir.FontBold = False
End Sub

Private Sub cmdguardar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdguardar.FontBold = True
Me.cmdmodificar.FontBold = False
Me.cmdcancelar.FontBold = False
Me.cmdsalir.FontBold = False
End Sub

Private Sub cmdcancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdcancelar.FontBold = True
Me.cmdguardar.FontBold = False
Me.cmdmodificar.FontBold = False
Me.cmdsalir.FontBold = False
End Sub

Private Sub cmdsalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdsalir.FontBold = True
Me.cmdguardar.FontBold = False
Me.cmdcancelar.FontBold = False
Me.cmdmodificar.FontBold = False
End Sub

Private Sub cmdcancelar_Click()

On Error GoTo ControlError
    
    instalacion.Recordset.CancelUpdate

            cmdguardar.Enabled = False
            cmdeliminar.Enabled = True
            cmdmodificar.Enabled = True
            cmdsalir.Enabled = True
            cmdcancelar.Visible = False
            Me.Command1(0).Enabled = True
            Me.Command1(1).Enabled = True
            Me.Command1(2).Enabled = True
            Me.Command1(3).Enabled = True
            Text1.Locked = True
            Text2.Locked = True
            Text4.Locked = True
            Text5.Locked = True
            Text8.Locked = True
            
    Exit Sub    ' Salir para evitar el controlador.

ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "Gestión y Control")
        
    End Select

 End Sub

Private Sub cmdmodificar_Click()
    
    cmdcancelar.Visible = True
    cmdsalir.Enabled = False
    cmdguardar.Enabled = True
    cmdmodificar.Enabled = False
 
       Command1(0).Enabled = False
       Command1(1).Enabled = False
       Command1(2).Enabled = False
       Command1(3).Enabled = False
        
    Text1.Locked = False
    Text2.Locked = False
    Text4.Locked = False
    Text5.Locked = False
    Text8.Locked = False
        
End Sub

Private Sub cmdguardar_Click()
Dim fec As Date
Dim ano As Date
Dim strquery As String
Dim bandera As Boolean
Dim abc, ncliente, contrato

On Error GoTo ControlError
    
    
    If IsNull(Text2.Text) Or Text1.Text = "" Then
    
        MsgBox "Debe Indicar el Precio de Venta, por favor verifique ", vbInformation, "JerGas"
             
        Me.Text2.SetFocus
             
        Exit Sub
    End If
    
  
 
    With instalacion.Recordset

        mvBookMark = .Bookmark

        .Update

        .Bookmark = mvBookMark
       instalacion.Refresh

    End With
    
'     If Text1.Text = "10" Then
'      With materiales.Recordset
'        mvBookMark = .Bookmark
'        .MoveFirst
'        .Find "id_inst = 10"
'           !precio = CInt(Me.Text2.Text)
'        .Update
'      End With
'     End If
'
'    If Text1.Text = "18" Then
'      With materiales.Recordset
'        mvBookMark = .Bookmark
'        .MoveFirst
'        .Find "id_inst = 18"
'           !precio = CInt(Me.Text2.Text)
'        .Update
'      End With
'    End If
'
'    If Text1.Text = "27" Then
'      With materiales.Recordset
'        mvBookMark = .Bookmark
'        .MoveFirst
'        .Find "id_inst = 27"
'           !precio = CInt(Me.Text2.Text)
'        .Update
'      End With
'    End If
'
'    If Text1.Text = "43" Then
'      With materiales.Recordset
'        mvBookMark = .Bookmark
'        .MoveFirst
'        .Find "id_inst = 43"
'           !precio = CInt(Me.Text2.Text)
'        .Update
'      End With
'    End If
    
    cmdmodificar.Enabled = True
    cmdsalir.Enabled = True
    cmdcancelar.Visible = True
    cmdguardar.Enabled = False
   
       Command1(0).Enabled = True
       Command1(1).Enabled = True
       Command1(2).Enabled = True
       Command1(3).Enabled = True

    
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
Me.Command1(0).Enabled = True
Me.Command1(1).Enabled = True
Me.Command1(2).Enabled = True
Me.Command1(3).Enabled = True

End Sub

Private Sub cmdsalir_Click()
  Unload Me
End Sub

Private Sub text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0

End Sub

