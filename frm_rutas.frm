VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm_rutas 
   Caption         =   "SISTEMA DE GESTIÓN Y CONTROL JER-GAS, C.A."
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10110
   Icon            =   "frm_rutas.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5655
   ScaleWidth      =   10110
   WindowState     =   2  'Maximized
   Begin VB.PictureBox txt_ruta5 
      Height          =   5295
      Left            =   6720
      ScaleHeight     =   5235
      ScaleWidth      =   7635
      TabIndex        =   23
      Top             =   1800
      Width           =   7695
   End
   Begin VB.PictureBox txt_ruta4 
      Height          =   5295
      Left            =   6720
      Picture         =   "frm_rutas.frx":08CA
      ScaleHeight     =   5235
      ScaleWidth      =   7635
      TabIndex        =   22
      Top             =   1800
      Width           =   7695
   End
   Begin VB.PictureBox txt_ruta3 
      Height          =   5295
      Left            =   6720
      Picture         =   "frm_rutas.frx":82CA4
      ScaleHeight     =   5235
      ScaleWidth      =   7635
      TabIndex        =   21
      Top             =   1800
      Width           =   7695
   End
   Begin VB.PictureBox txt_ruta2 
      Height          =   5295
      Left            =   6720
      Picture         =   "frm_rutas.frx":10507E
      ScaleHeight     =   5235
      ScaleWidth      =   7635
      TabIndex        =   20
      Top             =   1800
      Width           =   7695
   End
   Begin VB.PictureBox txt_ruta1 
      Height          =   5295
      Left            =   6720
      Picture         =   "frm_rutas.frx":187458
      ScaleHeight     =   5235
      ScaleWidth      =   7635
      TabIndex        =   19
      Top             =   1800
      Width           =   7695
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>|"
      Height          =   615
      Index           =   3
      Left            =   4080
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      Height          =   615
      Index           =   2
      Left            =   3360
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<<"
      Height          =   615
      Index           =   1
      Left            =   2640
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "|<<"
      Height          =   615
      Index           =   0
      Left            =   1920
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6480
      Width           =   735
   End
   Begin VB.Frame Frame13 
      Caption         =   "Zona"
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
      Left            =   2640
      TabIndex        =   12
      ToolTipText     =   "Suministre la Cédula o el Código del Usuario."
      Top             =   1920
      Width           =   1935
      Begin MSForms.TextBox Txt_zona 
         Bindings        =   "frm_rutas.frx":209832
         DataField       =   "zona"
         DataSource      =   "rutas"
         Height          =   300
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1695
         VariousPropertyBits=   746604571
         Size            =   "2990;529"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin VB.Frame Frame9 
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
      Height          =   1335
      Left            =   1200
      TabIndex        =   10
      ToolTipText     =   "Suministre el Nombre y Apellido del Usuario."
      Top             =   2640
      Width           =   4575
      Begin MSForms.TextBox Txt_descripcion 
         Bindings        =   "frm_rutas.frx":20985A
         DataField       =   "descripción"
         DataSource      =   "rutas"
         Height          =   1020
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4335
         VariousPropertyBits=   -1400879077
         Size            =   "7646;1799"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Id de Ruta"
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
      Left            =   1200
      TabIndex        =   8
      ToolTipText     =   "Suministre la Cédula o el Código del Usuario."
      Top             =   1920
      Width           =   1215
      Begin MSForms.TextBox Txt_id_ruta 
         Bindings        =   "frm_rutas.frx":209884
         DataField       =   "id_ruta"
         DataSource      =   "rutas"
         Height          =   300
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   615
         VariousPropertyBits=   746604571
         Size            =   "1085;529"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   1200
      TabIndex        =   0
      Top             =   4080
      Width           =   4575
      Begin VB.CommandButton cmdagregar 
         Caption         =   "&Agregar"
         Height          =   660
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Botón para Agregar un Nuevo Usuario"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdeliminar 
         Caption         =   "&Eliminar"
         Height          =   660
         Left            =   1200
         TabIndex        =   6
         ToolTipText     =   "Elimina de la Base de Datos a un Usuario"
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdguardar 
         Caption         =   "&Guardar"
         Height          =   660
         Left            =   2280
         TabIndex        =   5
         ToolTipText     =   "Para Salvar el Usuario Agregado o Modificado"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdsalir 
         Caption         =   "&Cerrar"
         Height          =   660
         Left            =   2280
         TabIndex        =   4
         ToolTipText     =   "Cerrar el Sistema"
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "&Cancelar"
         Height          =   660
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Pulse este botón si desea Cancelar el Usuario Agregado"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdcancelar 
         Caption         =   "&Cancelar"
         Height          =   660
         Left            =   3360
         TabIndex        =   2
         ToolTipText     =   "Busca un Usuario"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdmodificar 
         Caption         =   "&Modificar"
         Height          =   660
         Left            =   1200
         TabIndex        =   1
         ToolTipText     =   "Cambiar Característica de un Usuario"
         Top             =   240
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc rutas 
      Height          =   450
      Left            =   0
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
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   960
      Width           =   15705
   End
   Begin VB.Label Label_titulo 
      BackColor       =   &H80000001&
      Caption         =   "  CONTROL DE RUTAS"
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
      Left            =   6720
      TabIndex        =   14
      Top             =   240
      Width           =   8775
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000B&
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   0
      Top             =   900
      Width           =   15105
   End
End
Attribute VB_Name = "frm_rutas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdagregar_Click()
On Error GoTo AddErr
    
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
    
    txt_id_ruta.Locked = False
    txt_zona.Locked = False
    txt_descripcion.Locked = False
        
    With rutas.Recordset
    
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
            
        .MoveLast
    End If
    .AddNew
      
  End With
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdcancelar_Click()

On Error GoTo ControlError
    
    rutas.Recordset.CancelUpdate
            
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
            
                txt_id_ruta.Locked = True
                txt_zona.Locked = True
                txt_descripcion.Locked = True
         '       Txt_mapa.Locked = True

    Exit Sub    ' Salir para evitar el controlador.

ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "Control del rutases")
        
    End Select

 End Sub

Private Sub cmdeliminar_Click()
  'esto puede producir un error si elimina el último
  'registro o el único registro del recordset

On Error GoTo ControlError
    respuesta = MsgBox("¿Desea Eliminar el Registro?", vbYesNo)
    If respuesta = vbYes Then
        rutas.Recordset.Delete
        rutas.Recordset.MoveNext
    End If

    Exit Sub    ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "Control de rutases")
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

    txt_id_ruta.Locked = False
    txt_zona.Locked = False
    txt_descripcion.Locked = False
 '   Txt_mapa.Locked = False

End Sub

Private Sub cmdguardar_Click()
Dim fec As Date
Dim ano As Date
Dim strquery As String
Dim bandera As Boolean
Dim abc, ncliente, contrato

On Error GoTo ControlError
    
    
    If IsNull(txt_zona.Text) Or txt_zona.Text = "" Then
    
        MsgBox "Apellido y Nombre no puede ser nulo, por favor verifique ", vbInformation, "JerGas"
             
        Me.txt_zona.SetFocus
             
        Exit Sub
    End If
    
    If IsNull(Me.txt_id_ruta.Text) Or txt_id_ruta.Text = "" Then
    
        MsgBox "La ruta no puede ser nulo, por favor verifique ", vbInformation, "JerGas"
             
        Me.txt_id_ruta.SetFocus
             
        Exit Sub
    End If
    
    If IsNull(Me.txt_descripcion.Text) Or txt_descripcion.Text = "" Then
    
        MsgBox "Debe seleccionar un sector, por favor verifique ", vbInformation, "JerGas"
             
        Me.txt_descripcion.SetFocus
             
        Exit Sub
    End If
    
    With rutas.Recordset

        mvBookMark = .Bookmark

        .Update

        .Bookmark = mvBookMark

    End With
    
    cmdeliminar.Enabled = True
    cmdmodificar.Enabled = True
    cmdagregar.Visible = True
    cmdsalir.Enabled = True
    cmdcancelar.Visible = True
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
    cmdcancelar.Visible = True
Me.Command1(0).Enabled = True
Me.Command1(1).Enabled = True
Me.Command1(2).Enabled = True
Me.Command1(3).Enabled = True
  txt_ruta5.Visible = False
            
     txt_ruta1.Visible = True
     txt_ruta2.Visible = False
     txt_ruta3.Visible = False
     txt_ruta4.Visible = False

txt_id_ruta.Locked = True
txt_zona.Locked = True
txt_descripcion.Locked = True
End Sub

Private Sub cmdsalir_Click()
  Unload Me
End Sub

Private Sub Command1_Click(Index As Integer)
On Error GoTo ControlError

Select Case Index
    Case 0
       rutas.Recordset.MoveFirst
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
       rutas.Recordset.MovePrevious
       Command1(2).Enabled = True
       Command1(3).Enabled = True
         Me.cmdsalir.FontBold = False
         Me.cmdguardar.FontBold = False
         Me.cmdeliminar.FontBold = False
         Me.cmdcancelar.FontBold = False
         Me.cmdmodificar.FontBold = False
         Me.cmdagregar.FontBold = False
       If rutas.Recordset.BOF = True Then
        rutas.Recordset.MoveFirst
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
       rutas.Recordset.MoveNext
       Command1(0).Enabled = True
       Command1(1).Enabled = True
         Me.cmdsalir.FontBold = False
         Me.cmdguardar.FontBold = False
         Me.cmdeliminar.FontBold = False
         Me.cmdcancelar.FontBold = False
         Me.cmdmodificar.FontBold = False
         Me.cmdagregar.FontBold = False
       If rutas.Recordset.EOF = True Then
         Command1(2).Enabled = False
         Command1(3).Enabled = False
         Me.cmdsalir.FontBold = False
         Me.cmdguardar.FontBold = False
         Me.cmdeliminar.FontBold = False
         Me.cmdcancelar.FontBold = False
         Me.cmdmodificar.FontBold = False
         Me.cmdagregar.FontBold = False
         rutas.Recordset.MoveLast
       Else
       End If
    Case 3
       rutas.Recordset.MoveLast
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

Private Sub habilitar_botones(valor As Boolean)
'Me.cmd_rutas.Enabled = valor
'Me.cmd_estado.Enabled = valor
'Me.cmd_liquidacion.Enabled = valor

End Sub

Private Sub Txt_id_ruta_Change()
On Error GoTo ControlError

          If txt_id_ruta.Text = 1 Then
              txt_ruta1.Visible = True
              txt_ruta2.Visible = False
              txt_ruta3.Visible = False
              txt_ruta4.Visible = False
          End If

          If txt_id_ruta.Text = 2 Then
              txt_ruta1.Visible = False
              txt_ruta2.Visible = True
              txt_ruta3.Visible = False
              txt_ruta4.Visible = False
          End If

          If txt_id_ruta.Text = 3 Then
              txt_ruta1.Visible = False
              txt_ruta2.Visible = False
              txt_ruta3.Visible = True
              txt_ruta4.Visible = False
          End If
          
          If txt_id_ruta.Text = 4 Then
              txt_ruta1.Visible = False
              txt_ruta2.Visible = False
              txt_ruta3.Visible = False
              txt_ruta4.Visible = True
          
          End If

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
