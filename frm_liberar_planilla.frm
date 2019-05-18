VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_liberar_planilla 
   Caption         =   "Liberar una Factura"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6585
   Icon            =   "frm_liberar_planilla.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   6585
   Begin VB.TextBox txt_liquidado 
      DataField       =   "id_pedido"
      DataSource      =   "liquidado"
      Height          =   285
      Left            =   2760
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txt_id_pedidos 
      DataField       =   "id_pedido"
      DataSource      =   "pedidos"
      Height          =   285
      Left            =   480
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Introduzca el Número de Factura"
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6135
      Begin VB.CommandButton cmd_eliminar 
         Caption         =   "Eliminar"
         Height          =   615
         Left            =   3240
         TabIndex        =   8
         ToolTipText     =   "Elimina una Factura"
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "Cerrar"
         Height          =   615
         Left            =   4560
         TabIndex        =   3
         ToolTipText     =   "Cerrar la Pantalla"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmd_liberar 
         Caption         =   "Liberar"
         Height          =   615
         Left            =   3240
         TabIndex        =   2
         ToolTipText     =   "Libera una Factura"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txt_planilla 
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
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label lbl_informa_liq 
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Lbl_informacion 
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   2775
      End
   End
   Begin MSAdodcLib.Adodc pedidos 
      Height          =   375
      Left            =   480
      Top             =   2520
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "tbl_pedidos"
      Caption         =   "pedidos"
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
   Begin MSAdodcLib.Adodc liquidado 
      Height          =   375
      Left            =   2760
      Top             =   2520
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "tbl_liquidado"
      Caption         =   "liquidado"
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
Attribute VB_Name = "frm_liberar_planilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cerrar_Click()
Unload Me
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = True
Me.cmd_liberar.FontBold = False

End Sub

Private Sub cmd_eliminar_Click()
On Error GoTo ControlError
Dim Cod
Cod = InputBox("Suministre la clave para eliminar la planilla", "JerGas C.A.")

If Cod = "jergas1221" Then
    
    pedidos.Recordset.MoveFirst
    
    strquery = "id_pedido = '" & Me.txt_planilla.Text & "'"

    pedidos.Recordset.Find strquery
    
    If pedidos.Recordset.EOF Then
    
            MsgBox "Nºde Planilla suministrada no encontrada, por favor verifique ", vbInformation, "JerGas C.A."
            
            Me.Lbl_informacion.Caption = "Planilla No Encontrada"
            
            Exit Sub
                    
    Else
            
            pedidos.Recordset.Delete
            
            Me.Lbl_informacion.Caption = "Planilla Eliminada"
            
            'Liberando en liquidacion
            liquidado.Recordset.MoveFirst
            
            strquery = "id_pedido = '" & Me.txt_planilla.Text & "'"
        
            liquidado.Recordset.Find strquery
            
            If liquidado.Recordset.EOF Then
            
                    MsgBox "Nºde Planilla suministrada no encontrada, por favor verifique ", vbInformation, "JerGas C.A."
                    
                    Me.Lbl_informacion.Caption = "Planilla No Encontrada"
                    
                    Exit Sub
                            
            Else
                liquidado.Recordset.Delete
                lbl_informa_liq.Caption = "Planilla borrada de liquidación"
            End If
            
            liquidado.Recordset.Close

    End If
    
    pedidos.Recordset.Close
End If
Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "JerGas C.A.")
        
    End Select

End Sub

Private Sub cmd_liberar_Click()
On Error GoTo ControlError

    pedidos.Recordset.MoveFirst
    
    strquery = "id_pedido = '" & Me.txt_planilla.Text & "'"

    pedidos.Recordset.Find strquery
    
    If pedidos.Recordset.EOF Then
    
            MsgBox "Nºde Planilla suministrada no encontrada, por favor verifique ", vbInformation, "JerGas C.A."
            
            Me.Lbl_informacion.Caption = "Planilla No Encontrada"
            
            Exit Sub
                    
    Else
    
            
            
            pedidos.Recordset!status = "VI"
                
            pedidos.Recordset!fecha_cancel = Null
            
            pedidos.Recordset.Update
            
            Me.Lbl_informacion.Caption = "Planilla Liberada"
            
            'Liberando en liquidacion
            liquidado.Recordset.MoveFirst
            
            strquery = "id_pedido = '" & Me.txt_planilla.Text & "'"
        
            liquidado.Recordset.Find strquery
            
            If liquidado.Recordset.EOF Then
            
                    'MsgBox "Nºde Planilla suministrada no encontrada, por favor verifique ", vbInformation, "JerGas C.A."
                    
                    Me.Lbl_informacion.Caption = "Planilla No Encontrada"
                    
                    Exit Sub
                            
            Else
                liquidado.Recordset.Delete
                lbl_informa_liq.Caption = "Planilla borrada de liquidación"
            End If
            
            liquidado.Recordset.Close

    End If
    
    pedidos.Recordset.Close
Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "JerGas C.A.")
        
    End Select
End Sub

Private Sub cmd_liberar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_liberar.FontBold = True
End Sub

Private Sub Form_Load()
Me.Height = 2700
Me.Width = 6700
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_liberar.FontBold = False

End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_liberar.FontBold = False

End Sub

Private Sub txt_planilla_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub
