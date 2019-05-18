VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_choferes 
   Caption         =   "SISTEMA DE GESTIÓN Y CONTROL JER-GAS, C.A."
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11160
   Icon            =   "frm_choferes2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   ">>|"
      Height          =   615
      Index           =   3
      Left            =   10920
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      Height          =   615
      Index           =   2
      Left            =   10200
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<<"
      Height          =   615
      Index           =   1
      Left            =   9480
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "|<<"
      Height          =   615
      Index           =   0
      Left            =   8760
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4560
      Width           =   735
   End
   Begin VB.Frame Frame17 
      Height          =   1215
      Left            =   4680
      TabIndex        =   18
      Top             =   5280
      Width           =   6495
      Begin VB.CommandButton cmdagregar 
         Caption         =   "&Agregar"
         Height          =   660
         Left            =   240
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Botón para Agregar un Nuevo Usuario"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdcancelar 
         Caption         =   "&Cancelar"
         Height          =   660
         Left            =   240
         TabIndex        =   23
         ToolTipText     =   "Pulse este botón si desea Cancelar el Usuario Agregado"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdguardar 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   660
         Left            =   1440
         TabIndex        =   22
         ToolTipText     =   "Para Salvar el Usuario Agregado o Modificado"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdmodificar 
         Caption         =   "&Modificar"
         Height          =   660
         Left            =   2640
         TabIndex        =   21
         ToolTipText     =   "Cambiar Característica de un Usuario"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdeliminar 
         Caption         =   "&Eliminar"
         Height          =   660
         Left            =   3840
         TabIndex        =   20
         ToolTipText     =   "Elimina de la Base de Datos a un Usuario"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdsalir 
         Caption         =   "C&errar"
         Height          =   660
         Left            =   5040
         TabIndex        =   19
         ToolTipText     =   "Cerrar el Sistema"
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc chofer 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      RecordSource    =   "tbl_chofer"
      Caption         =   "Adodc1"
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
   Begin VB.Frame Frame5 
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
      Height          =   650
      Left            =   7440
      TabIndex        =   16
      ToolTipText     =   "Suministre la Cédula o el Código del Usuario."
      Top             =   1800
      Width           =   1935
      Begin VB.TextBox txt_sector 
         Alignment       =   2  'Center
         DataField       =   "zona"
         DataSource      =   "chofer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Código de Chofer"
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
      Height          =   650
      Left            =   4440
      TabIndex        =   15
      ToolTipText     =   "Suministre la Cédula o el Código del Usuario."
      Top             =   1800
      Width           =   1695
      Begin VB.TextBox txt_codigo 
         Alignment       =   2  'Center
         DataField       =   "id_chofer"
         DataSource      =   "chofer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Teléfonos"
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
      Height          =   930
      Left            =   4440
      TabIndex        =   12
      ToolTipText     =   "Suministre el Teléfono del Usuario."
      Top             =   4320
      Width           =   3975
      Begin VB.TextBox txt_telefono_cel 
         Alignment       =   2  'Center
         DataField       =   "telefono_cel"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "(####) ### ## ##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "chofer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txt_telefono_hab 
         Alignment       =   2  'Center
         DataField       =   "telefono_hab"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "(####) ### ## ##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "chofer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Habitación"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Celular"
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Apellidos y Nombres"
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
      Height          =   650
      Left            =   4440
      TabIndex        =   11
      ToolTipText     =   "Suministre el Nombre y Apellido del Usuario."
      Top             =   2640
      Width           =   3615
      Begin VB.TextBox txt_nombre 
         DataField       =   "nombre"
         DataSource      =   "chofer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Cédula de Identidad"
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
      Height          =   650
      Left            =   8280
      TabIndex        =   10
      ToolTipText     =   "Suministre la Cédula o el Código del Usuario."
      Top             =   2640
      Width           =   2055
      Begin VB.TextBox txt_cedula 
         Alignment       =   2  'Center
         DataField       =   "cedula"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "chofer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Dirección"
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
      Height          =   650
      Left            =   4440
      TabIndex        =   9
      ToolTipText     =   "Suministre el Nombre y Apellido del Usuario."
      Top             =   3480
      Width           =   5895
      Begin VB.TextBox txt_direccion 
         DataField       =   "direccion"
         DataSource      =   "chofer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "Ruta"
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
      Height          =   650
      Left            =   6360
      TabIndex        =   8
      ToolTipText     =   "Suministre la Cédula o el Código del Usuario."
      Top             =   1800
      Width           =   855
      Begin VB.TextBox txt_ruta 
         Alignment       =   2  'Center
         DataField       =   "id_ruta"
         DataSource      =   "chofer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000000&
      BorderColor     =   &H8000000B&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   960
      Width           =   15465
   End
   Begin VB.Label Label_titulo 
      BackColor       =   &H80000001&
      Caption         =   "  CONTROL DE CONDUCTORES"
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
      Left            =   6600
      TabIndex        =   17
      Top             =   240
      Width           =   8655
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000000&
      BorderColor     =   &H8000000B&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   690
      Left            =   0
      Top             =   900
      Width           =   15465
   End
End
Attribute VB_Name = "frm_choferes"
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
    
    txt_codigo.Locked = False
    txt_ruta.Locked = False
    txt_sector.Locked = False
    txt_nombre.Locked = False
    txt_cedula.Locked = False
    txt_direccion.Locked = False
    txt_telefono_hab.Locked = False
    txt_telefono_cel.Locked = False
    
    
    With chofer.Recordset
    
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
    
    chofer.Recordset.CancelUpdate
            
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
            
                txt_codigo.Locked = True
                txt_ruta.Locked = True
                txt_sector.Locked = True
                txt_nombre.Locked = True
                txt_cedula.Locked = True
                txt_direccion.Locked = True
                txt_telefono_hab.Locked = True
                txt_telefono_cel.Locked = True

    Exit Sub    ' Salir para evitar el controlador.

ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "Control del Choferes")
        
    End Select

 End Sub

Private Sub cmdeliminar_Click()
  'esto puede producir un error si elimina el último
  'registro o el único registro del recordset

On Error GoTo ControlError
    respuesta = MsgBox("¿Desea Eliminar el Registro?", vbYesNo)
    If respuesta = vbYes Then
        chofer.Recordset.Delete
        chofer.Recordset.MoveNext
    End If

    Exit Sub    ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "Control de Choferes")
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

    txt_codigo.Locked = False
    txt_ruta.Locked = False
    txt_sector.Locked = False
    txt_nombre.Locked = False
    txt_cedula.Locked = False
    txt_direccion.Locked = False
    txt_telefono_hab.Locked = False
    txt_telefono_cel.Locked = False

End Sub

Private Sub cmdguardar_Click()
Dim fec As Date
Dim ano As Date
Dim strquery As String
Dim bandera As Boolean
Dim abc, ncliente, contrato

On Error GoTo ControlError
    
    
    If IsNull(txt_nombre.Text) Or txt_nombre.Text = "" Then
    
        MsgBox "Apellido y Nombre no puede ser nulo, por favor verifique ", vbInformation, "JerGas C.A."
             
        Me.txt_nombre.SetFocus
             
        Exit Sub
    End If
    
    If IsNull(Me.txt_ruta.Text) Or txt_ruta.Text = "" Then
    
        MsgBox "La ruta no puede ser nulo, por favor verifique ", vbInformation, "JerGas C.A."
             
        Me.txt_ruta.SetFocus
             
        Exit Sub
    End If
    
    If IsNull(Me.txt_sector.Text) Or txt_sector.Text = "" Then
    
        MsgBox "Debe seleccionar un sector, por favor verifique ", vbInformation, "JerGas C.A."
             
        Me.txt_sector.SetFocus
             
        Exit Sub
    End If
    
    With chofer.Recordset

        mvBookMark = .Bookmark

        .Update

        .Bookmark = mvBookMark

    End With
    
    cmdeliminar.Enabled = True
    cmdmodificar.Enabled = True
    cmdagregar.Visible = True
    cmdsalir.Enabled = True
    cmdcancelar.Visible = False
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
    cmdcancelar.Visible = False
Me.Command1(0).Enabled = True
Me.Command1(1).Enabled = True
Me.Command1(2).Enabled = True
Me.Command1(3).Enabled = True

txt_codigo.Locked = True
txt_ruta.Locked = True
txt_sector.Locked = True
txt_nombre.Locked = True
txt_cedula.Locked = True
txt_direccion.Locked = True
txt_telefono_hab.Locked = True
txt_telefono_cel.Locked = True
End Sub

Private Sub cmdsalir_Click()
  Unload Me
End Sub

Private Sub Command1_Click(Index As Integer)
On Error GoTo ControlError

Select Case Index
    Case 0
       chofer.Recordset.MoveFirst
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
       chofer.Recordset.MovePrevious
       Command1(2).Enabled = True
       Command1(3).Enabled = True
         Me.cmdsalir.FontBold = False
         Me.cmdguardar.FontBold = False
         Me.cmdeliminar.FontBold = False
         Me.cmdcancelar.FontBold = False
         Me.cmdmodificar.FontBold = False
         Me.cmdagregar.FontBold = False
       If chofer.Recordset.BOF = True Then
        chofer.Recordset.MoveFirst
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
       chofer.Recordset.MoveNext
       Command1(0).Enabled = True
       Command1(1).Enabled = True
         Me.cmdsalir.FontBold = False
         Me.cmdguardar.FontBold = False
         Me.cmdeliminar.FontBold = False
         Me.cmdcancelar.FontBold = False
         Me.cmdmodificar.FontBold = False
         Me.cmdagregar.FontBold = False
       If chofer.Recordset.EOF = True Then
         Command1(2).Enabled = False
         Command1(3).Enabled = False
         Me.cmdsalir.FontBold = False
         Me.cmdguardar.FontBold = False
         Me.cmdeliminar.FontBold = False
         Me.cmdcancelar.FontBold = False
         Me.cmdmodificar.FontBold = False
         Me.cmdagregar.FontBold = False
         chofer.Recordset.MoveLast
       Else
       End If
    Case 3
       chofer.Recordset.MoveLast
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
'Me.cmd_chofer.Enabled = valor
'Me.cmd_estado.Enabled = valor
'Me.cmd_liquidacion.Enabled = valor

End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_ruta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_sector_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_nombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Index <> 6 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 13 Or Index = 6 Or Index = 9 Then
        If ((KeyAscii < 48) Or (KeyAscii > 57)) And (KeyAscii <> 8) Then KeyAscii = 0
    End If
End Sub

Private Sub txt_cedula_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_direccion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Index <> 6 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 13 Or Index = 6 Or Index = 9 Then
        If ((KeyAscii < 48) Or (KeyAscii > 57)) And (KeyAscii <> 8) Then KeyAscii = 0
    End If
End Sub

Private Sub txt_telefono_cel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_telefono_hab_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub



