VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_historico_pedidos_clientes 
   Caption         =   "SISTEMA DE GESTIÓN Y CONTROL "
   ClientHeight    =   2100
   ClientLeft      =   3555
   ClientTop       =   3945
   ClientWidth     =   8205
   Icon            =   "frm_historico_pedidos_clientes.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   8205
   Begin VB.CommandButton cmd_cerrar 
      Caption         =   "Cerrar"
      Height          =   615
      Left            =   6720
      TabIndex        =   10
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmd_procesar 
      Caption         =   "Procesar"
      Height          =   615
      Left            =   5280
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      DataField       =   "contrato_num"
      DataSource      =   "clientes"
      Height          =   285
      Left            =   5520
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      DataField       =   "direccion"
      DataSource      =   "clientes"
      Height          =   285
      Left            =   4200
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      DataField       =   "cedula"
      DataSource      =   "clientes"
      Height          =   285
      Left            =   2880
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txt_clientes 
      DataField       =   "cliente"
      DataSource      =   "clientes"
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "codigo"
      DataSource      =   "clientes"
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   8415
      Begin VB.CommandButton Busquedad_avanzadas 
         Caption         =   "Búsqueda Avanzada"
         Height          =   375
         Index           =   0
         Left            =   6000
         TabIndex        =   1
         Tag             =   "Lista todos los inmuebles registrados"
         Top             =   120
         Width           =   1935
      End
      Begin MSDataListLib.DataCombo Dcmb_Buscar 
         Bindings        =   "frm_historico_pedidos_clientes.frx":08CA
         Height          =   315
         Left            =   480
         TabIndex        =   2
         ToolTipText     =   "Pulse doble click para cambiar el tipo de busqueda, después de presionar búsqueda avanzada"
         Top             =   120
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         ListField       =   "codigo"
         BoundColumn     =   "codigo"
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc clientes 
      Height          =   375
      Left            =   3840
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
      RecordSource    =   "tbl_clientes"
      Caption         =   "clientes"
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
   Begin VB.Label LABEL_BUSCA 
      BackStyle       =   0  'Transparent
      Caption         =   "Búsqueda por Código: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   3615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000B&
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   0
      Top             =   540
      Width           =   15345
   End
End
Attribute VB_Name = "frm_historico_pedidos_clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Busquedad_avanzadas_Click(Index As Integer)
            Busq_Avanzada = True
            clientes.CommandType = adCmdText
            clientes.RecordSource = "select * from tbl_clientes WHERE codigo <> '' ORDER BY codigo ASC"
            clientes.Refresh
            Call Dcmb_Buscar_Click(1)
End Sub

Private Sub buscar_codigo()

On Error GoTo ControlError

Dim strquery, RESP

If Not Busq_Avanzada And ((Dcmb_Buscar.Text Like "%*%" Or Dcmb_Buscar.Text Like "%*" Or Dcmb_Buscar.Text Like "*%") Or (Me.clientes.Recordset.RecordCount = 0)) Then
    
    Me.clientes.CommandType = adCmdText
    
    Me.clientes.RecordSource = "SELECT * FROM tbl_clientes WHERE tbl_clientes.codigo like '" & Dcmb_Buscar.Text & "' ORDER BY codigo"
    
    Me.clientes.Refresh

    If clientes.Recordset.EOF Then
    
        MsgBox "Código suministrado no encontrado, por favor verifique ", vbInformation, "JerGas C.A."
        
        Dcmb_Buscar.Text = ""
        
        Dcmb_Buscar.SetFocus
        
    Else
    
        If Me.clientes.Recordset.RecordCount > 1 Then
        
            MsgBox Me.clientes.Recordset.RecordCount & " encontrados"
            
            Busq_Avanzada = True
            
        End If
    End If
    
Else
    
    clientes.Recordset.MoveFirst
    
    strquery = "codigo = '" & Dcmb_Buscar.Text & "'"

    clientes.Recordset.Find strquery
    
    If clientes.Recordset.EOF Then
    
            MsgBox "Código suministrado no encontrado, por favor verifique ", vbInformation, "SIAGEP"
            
            Dcmb_Buscar.Text = ""
            
            Dcmb_Buscar.SetFocus
      Else
    End If
    

End If

Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "JerGas")
        Case 3001
            v = MsgBox("Nombre suministrado no encontrado", vbOKOnly, "JerGas")
    End Select

End Sub

Private Sub Dcmb_Buscar_Click(Area As Integer)
If Area = 2 Then
        If Dcmb_Buscar.ListField = "codigo" Then
            If Dcmb_Buscar.Text <> "" Then
                
                Call buscar_codigo
                Call buscar_pedidos
            Else
                Exit Sub
            End If
        End If
        
        If Dcmb_Buscar.ListField = "contrato_num" Then
            If Dcmb_Buscar.Text <> "" Then
                Call buscar_contrato_num
                Call buscar_pedidos
            Else
                Exit Sub
            End If
        End If

        If Dcmb_Buscar.ListField = "cliente" Then
            If Dcmb_Buscar.Text <> "" Then
                Call buscar_cliente
                Call buscar_pedidos
            Else
                Exit Sub
            End If
        End If

        If Dcmb_Buscar.ListField = "cedula" Then
            If Dcmb_Buscar.Text <> "" Then
                Call buscar_cedula
                Call buscar_pedidos
            Else
                Exit Sub
            End If
        End If
        If Dcmb_Buscar.ListField = "direccion" Then
            If Dcmb_Buscar.Text <> "" Then
                Call buscar_direccion
                Call buscar_pedidos
            Else
                Exit Sub
            End If
        End If
End If
End Sub
Private Sub buscar_contrato_num()

On Error GoTo ControlError

Dim strquery, RESP

If Not Busq_Avanzada And ((Dcmb_Buscar.Text Like "%*%" Or Dcmb_Buscar.Text Like "%*" Or Dcmb_Buscar.Text Like "*%") Or (Me.clientes.Recordset.RecordCount = 0)) Then
    
    Me.clientes.CommandType = adCmdText
    
    Me.clientes.RecordSource = "SELECT * FROM tbl_clientes WHERE tbl_clientes.contrato_num like '" & Dcmb_Buscar.Text & "' ORDER BY contrato_num"
    
    Me.clientes.Refresh

    If clientes.Recordset.EOF Then
    
        MsgBox "Nº de Contrato suministrado no encontrado, por favor verifique ", vbInformation, "JerGas C.A."
        
        Dcmb_Buscar.Text = ""
        
        Dcmb_Buscar.SetFocus
    Else
    
        If Me.clientes.Recordset.RecordCount > 1 Then
        
            MsgBox Me.clientes.Recordset.RecordCount & " encontrados"
            
            Busq_Avanzada = True
            
        End If
    End If
    
Else
    
    clientes.Recordset.MoveFirst
    
    strquery = "contrato_num = '" & Dcmb_Buscar.Text & "'"

    clientes.Recordset.Find strquery
    
    If clientes.Recordset.EOF Then
    
            MsgBox "Nº de Contrato suministrado no encontrado, por favor verifique ", vbInformation, "SIAGEP"
            
            Dcmb_Buscar.Text = ""
            
            Dcmb_Buscar.SetFocus
    Else
    End If
    

End If

Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "JerGas")
        Case 3001
            v = MsgBox("Contrato suministrado no encontrado", vbOKOnly, "JerGas")
    End Select

End Sub
Private Sub buscar_cedula()

On Error GoTo ControlError

Dim strquery, RESP

If Not Busq_Avanzada And ((Dcmb_Buscar.Text Like "%*%" Or Dcmb_Buscar.Text Like "%*" Or Dcmb_Buscar.Text Like "*%") Or (Me.clientes.Recordset.RecordCount = 0)) Then
    
    Me.clientes.CommandType = adCmdText
    
    Me.clientes.RecordSource = "SELECT * FROM tbl_clientes WHERE tbl_clientes.cedula like '" & Dcmb_Buscar.Text & "' ORDER BY cedula"
    
    Me.clientes.Refresh

    If clientes.Recordset.EOF Then
    
        MsgBox "Cédula del cliente suministrado no encontrado, por favor verifique ", vbInformation, "JerGas C.A."
        
        Dcmb_Buscar.Text = ""
        
        Dcmb_Buscar.SetFocus
    Else
    
        If Me.clientes.Recordset.RecordCount > 1 Then
        
            MsgBox Me.clientes.Recordset.RecordCount & " encontrados"
            
            Busq_Avanzada = True
            
        End If
    End If
    
Else
    
    clientes.Recordset.MoveFirst
    
    strquery = "cedula = '" & Dcmb_Buscar.Text & "'"

    clientes.Recordset.Find strquery
    
    If clientes.Recordset.EOF Then
    
            MsgBox "Cédula del Cliente suministrado no encontrado, por favor verifique ", vbInformation, "SIAGEP"
            
            Dcmb_Buscar.Text = ""
            
            Dcmb_Buscar.SetFocus
    Else
    End If
    

End If

Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "JerGas")
        Case 3001
            v = MsgBox("Cédula del Cliente suministrado no encontrado", vbOKOnly, "JerGas")
    End Select

End Sub

Private Sub buscar_cliente()

On Error GoTo ControlError

Dim strquery, RESP

If Not Busq_Avanzada And ((Dcmb_Buscar.Text Like "%*%" Or Dcmb_Buscar.Text Like "%*" Or Dcmb_Buscar.Text Like "*%") Or (Me.clientes.Recordset.RecordCount = 0)) Then
    
    Me.clientes.CommandType = adCmdText
    
    Me.clientes.RecordSource = "SELECT * FROM tbl_clientes WHERE tbl_clientes.cliente like '" & Dcmb_Buscar.Text & "' ORDER BY cliente"
    
    Me.clientes.Refresh

    If clientes.Recordset.EOF Then
    
        MsgBox "Nombre del cliente suministrado no encontrado, por favor verifique ", vbInformation, "JerGas C.A."
        
        Dcmb_Buscar.Text = ""
        
        Dcmb_Buscar.SetFocus
    Else
    
        If Me.clientes.Recordset.RecordCount > 1 Then
        
            MsgBox Me.clientes.Recordset.RecordCount & " encontrados"
            
            Busq_Avanzada = True
            
        End If
    End If
    
Else
    
    clientes.Recordset.MoveFirst
    
    strquery = "cliente = '" & Dcmb_Buscar.Text & "'"

    clientes.Recordset.Find strquery
    
    If clientes.Recordset.EOF Then
    
            MsgBox "Nombre del Cliente suministrado no encontrado, por favor verifique ", vbInformation, "JerGas"
            
            Dcmb_Buscar.Text = ""
            
            Dcmb_Buscar.SetFocus
    Else
    End If
    

End If

Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "JerGas")
        Case 3001
            v = MsgBox("Nombre del Cliente suministrado no encontrado", vbOKOnly, "JerGas")
    End Select

End Sub

Private Sub buscar_direccion()

On Error GoTo ControlError

Dim strquery, RESP

If Not Busq_Avanzada And ((Dcmb_Buscar.Text Like "%*%" Or Dcmb_Buscar.Text Like "%*" Or Dcmb_Buscar.Text Like "*%") Or (Me.clientes.Recordset.RecordCount = 0)) Then
    
    Me.clientes.CommandType = adCmdText
    
    Me.clientes.RecordSource = "SELECT * FROM tbl_clientes WHERE tbl_clientes.direccion like '" & Dcmb_Buscar.Text & "' ORDER BY direccion"
    
    Me.clientes.Refresh

    If clientes.Recordset.EOF Then
    
        MsgBox "Direccion del cliente suministrada no encontrada, por favor verifique ", vbInformation, "JerGas C.A."
        
        Dcmb_Buscar.Text = ""
        
        Dcmb_Buscar.SetFocus
    Else
    
        If Me.clientes.Recordset.RecordCount > 1 Then
        
            MsgBox Me.clientes.Recordset.RecordCount & " encontrados"
            
            Busq_Avanzada = True
            
        End If
    End If
    
Else
    
    clientes.Recordset.MoveFirst
    
    strquery = "direccion = '" & Dcmb_Buscar.Text & "'"

    clientes.Recordset.Find strquery
    
    If clientes.Recordset.EOF Then
    
            MsgBox "Dirección del Cliente suministrado no encontrado, por favor verifique ", vbInformation, "JerGas"
            
            Dcmb_Buscar.Text = ""
            
            Dcmb_Buscar.SetFocus
    Else
    End If
    

End If

Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "JerGas")
        Case 3001
            v = MsgBox("Dirección del Cliente suministrado no encontrado", vbOKOnly, "JerGas")
    End Select

End Sub

Private Sub Dcmb_Buscar_DblClick(Area As Integer)
'Esta funci[on redefine el tipo de busqueda
'If Busq_Avanzada Then

    Me.Dcmb_Buscar.ToolTipText = "Doble click para alternar el tipo de busqueda"
    
    If Dcmb_Buscar.ListField = "codigo" Then
        'Si es bif pasa a ape
        If Busq_Avanzada Then
        
            Me.clientes.CommandType = adCmdText

            Me.clientes.RecordSource = "select * from tbl_clientes WHERE tbl_clientes.contrato_num <> '' ORDER BY contrato_num ASC"

            Me.clientes.Refresh
            
        End If

        Dcmb_Buscar.ListField = "contrato_num"

        Dcmb_Buscar.Text = ""

        LABEL_BUSCA.Caption = "Búsqueda por Nº de Contrato:"

        Exit Sub
        
    End If

    If Dcmb_Buscar.ListField = "contrato_num" Then
    
        'Si es ape pasa a cod cata
        If Busq_Avanzada Then
        
            Me.clientes.CommandType = adCmdText

            Me.clientes.RecordSource = "select * from tbl_clientes WHERE tbl_clientes.cliente <> '' ORDER BY cliente ASC"

            Me.clientes.Refresh
            
        End If
        
        Dcmb_Buscar.ListField = "cliente"

        Dcmb_Buscar.Text = ""

        LABEL_BUSCA.Caption = "Búsqueda por Cliente:"

        Exit Sub
        
    End If

    If Dcmb_Buscar.ListField = "cliente" Then

        'Si es cod pasa a cedula
        If Busq_Avanzada Then
        
            clientes.CommandType = adCmdText

            clientes.RecordSource = "select * from tbl_clientes WHERE tbl_clientes.cedula <> '' ORDER BY cedula ASC"

            clientes.Refresh
            
        End If
        
        Dcmb_Buscar.ListField = "cedula"

        Dcmb_Buscar.Text = ""

        LABEL_BUSCA.Caption = "Búsqueda por Cédula: "

        Exit Sub

    End If

    If Dcmb_Buscar.ListField = "cedula" Then
    
        'Si es cedual pasa a direccion
        If Busq_Avanzada Then
        
            clientes.CommandType = adCmdText

            clientes.RecordSource = "select * from tbl_clientes WHERE tbl_clientes.direccion <> '' ORDER BY direccion ASC"

            clientes.Refresh
            
        End If
        
        Dcmb_Buscar.ListField = "direccion"

        Dcmb_Buscar.Text = ""

        LABEL_BUSCA.Caption = "Búsqueda por Dirección: "

        Exit Sub
        
    End If

    If Dcmb_Buscar.ListField = "direccion" Then

        'Si es direccion pasa a bif
        If Busq_Avanzada Then
        
            clientes.CommandType = adCmdText

            clientes.RecordSource = "select * from tbl_clientes WHERE codigo <> '' ORDER BY codigo ASC"

            clientes.Refresh
            
        End If
        
        Dcmb_Buscar.ListField = "codigo"

        Dcmb_Buscar.Text = ""

        LABEL_BUSCA.Caption = "Búsqueda por Código: "

        Exit Sub
    End If
'End If
End Sub

Private Sub Form_Load()
Call actualizar_cn("SQL Server")
'Me.txt_fecha_pedido = Date
'Me.txt_fecha_entrega = DateAdd("d", 1, Date)
'Me.txt_date = Date
End Sub

Private Sub cmd_cerrar_Click()
Unload Me
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = True
Me.cmd_procesar.FontBold = False
End Sub

Private Sub cmd_procesar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_procesar.FontBold = True
End Sub

Private Sub cmd_procesar_Click()
rpt_clientes6.Show
End Sub

Private Sub buscar_pedidos()

On Error GoTo ControlError

hist_pedidos.CommandType = adCmdText

hist_pedidos.RecordSource = "SELECT * FROM tbl_pedidos WHERE tbl_pedidos.codigo = '" & Me.txt_clientes.Text & "' ORDER BY fecha_pedido DESC"

hist_pedidos.Refresh

Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "JerGas")
        Case 3001
            v = MsgBox("Nombre suministrado no encontrado", vbOKOnly, "JerGas")
    End Select
    
End Sub
