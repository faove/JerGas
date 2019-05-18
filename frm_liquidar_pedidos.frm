VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_liquidar_pedidos 
   Caption         =   "SISTEMA DE GESTIÓN Y CONTROL JER-GAS, C.A."
   ClientHeight    =   9210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "frm_liquidar_pedidos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9210
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_monto 
      Height          =   375
      Left            =   8040
      TabIndex        =   15
      Top             =   8400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt_fecha 
      Height          =   375
      Left            =   5400
      TabIndex        =   14
      Top             =   8400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt_cant 
      Height          =   375
      Left            =   7320
      TabIndex        =   13
      Top             =   8400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txt_inst 
      Height          =   375
      Left            =   6600
      TabIndex        =   12
      Top             =   8400
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSAdodcLib.Adodc resumen_mensual_ventas 
      Height          =   375
      Left            =   5760
      Top             =   0
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
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
      RecordSource    =   "tbl_resumen_mensual_ventas"
      Caption         =   "resumen_mensual_ventas"
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
   Begin VB.TextBox Text1 
      DataField       =   "fecha_venta"
      DataSource      =   "resumen_mensual_ventas"
      Height          =   375
      Left            =   9240
      TabIndex        =   11
      Top             =   8400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txt_codigo 
      DataField       =   "codigo"
      DataSource      =   "res_inventario"
      Height          =   375
      Left            =   10320
      TabIndex        =   10
      Top             =   8520
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSAdodcLib.Adodc res_inventario 
      Height          =   375
      Left            =   4560
      Top             =   360
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      RecordSource    =   "tbl_resumen_inventario"
      Caption         =   "res_inventatio"
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
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   14040
      TabIndex        =   9
      Top             =   8520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      DataField       =   "cant_actual"
      DataSource      =   "materiales"
      Height          =   375
      Left            =   13080
      TabIndex        =   8
      Top             =   8520
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSAdodcLib.Adodc materiales 
      Height          =   375
      Left            =   2520
      Top             =   360
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
   Begin VB.TextBox Text3 
      DataField       =   "cilindro"
      DataSource      =   "inventario"
      Height          =   375
      Left            =   12000
      TabIndex        =   7
      Top             =   8520
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSAdodcLib.Adodc inventario 
      Height          =   375
      Left            =   120
      Top             =   360
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      RecordSource    =   "tbl_inventario"
      Caption         =   "inventario"
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
   Begin VB.TextBox txt_cant_pedido 
      DataField       =   "cant_pedido"
      DataSource      =   "hist_pedidos"
      Height          =   375
      Left            =   11400
      TabIndex        =   6
      Top             =   8520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt_id_inst 
      DataField       =   "id_inst"
      DataSource      =   "hist_pedidos"
      Height          =   375
      Left            =   10800
      TabIndex        =   5
      Top             =   8520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmd_cerrar 
      Caption         =   "Cerrar"
      Height          =   735
      Left            =   2640
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   8400
      Width           =   1455
   End
   Begin VB.Frame Frame5 
      Caption         =   "FACTURAS PENDIENTES POR LIQUIDACIÓN"
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
      Height          =   6495
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   15015
      Begin MSDataGridLib.DataGrid DGrid_pedidos 
         Bindings        =   "frm_liquidar_pedidos.frx":08CA
         Height          =   6015
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   360
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   10610
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "cliente"
            Caption         =   "cliente"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "cedula"
            Caption         =   "cedula"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "id_pedido"
            Caption         =   "id_pedido"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "codigo"
            Caption         =   "codigo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "status"
            Caption         =   "status"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "fecha_pedido"
            Caption         =   "fecha_pedido"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "fecha_cancel"
            Caption         =   "fecha_cancel"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "id_inst"
            Caption         =   "id_inst"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "monto_fac"
            Caption         =   "monto_fac"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "cant_pedido"
            Caption         =   "cant_pedido"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "usuario_liq"
            Caption         =   "usuario_liq"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   540,284
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   764,787
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmd_liquidacion 
      Caption         =   "Liquidar Factura"
      Height          =   735
      Left            =   1200
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   8400
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc liquidado 
      Height          =   375
      Left            =   120
      Top             =   0
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
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
   Begin MSAdodcLib.Adodc hist_pedidos 
      Height          =   375
      Left            =   3600
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      RecordSource    =   "liquidacion"
      Caption         =   "inventario"
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
      Caption         =   "  LIQUIDACIÓN DE FACTURAS"
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
      Left            =   7560
      TabIndex        =   3
      Top             =   240
      Width           =   7695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   0
      Top             =   960
      Width           =   15465
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000B&
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   0
      Top             =   900
      Width           =   15465
   End
End
Attribute VB_Name = "frm_liquidar_pedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fechando

Private Sub DGrid_historico_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_liquidacion.FontBold = False
End Sub

Private Sub DGrid_pedidos_Click()
    Me.cmd_liquidacion.Enabled = True
    For Each Var In DGrid_pedidos.SelBookmarks
    DGrid_pedidos.Bookmark = Var
    
    Me.txt_fecha.Text = Date
    Me.txt_inst.Text = DGrid_pedidos.Columns(7)
    Me.txt_cant.Text = DGrid_pedidos.Columns(9)
    Me.txt_monto.Text = DGrid_pedidos.Columns(8)
    fechando = DGrid_pedidos.Columns(5)
    
    'Si status es CA
    '---------------
    If DGrid_pedidos.Columns(1) = "CA" Then
            MsgBox "Factura ya está cancelada", vbInformation, "JerGas"
            DGrid_pedidos.SelBookmarks.Remove (DGrid_pedidos.SelBookmarks.Count - 1)
'            If DGrid_inm_liq.SelBookmarks.Count = 0 Then
'                Tex_Cuotas.Text = ""
'                Tex_Monto.Text = ""
'            End If
            Exit For
            
    End If
    Next
End Sub

Private Sub Form_Load()
Call actualizar_cn("SQL Server")
'Me.Text1.Text = Date
End Sub

Private Sub cmd_cerrar_Click()
Unload Me
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = True
Me.cmd_liquidacion.FontBold = False
End Sub

Private Sub cmd_clientes_Click()
Screen.MousePointer = 13
frm_pedidos.Show
Screen.MousePointer = 0
End Sub

Private Sub cmd_clientes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_liquidacion.FontBold = False
End Sub

Private Sub cmd_estado_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_liquidacion.FontBold = False
End Sub

Private Sub cmd_liquidacion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_liquidacion.FontBold = True
End Sub

Private Sub cmd_liquidacion_Click()
Dim actualiza As Integer
On Error GoTo control_error

'Desabilita el botón de aceptar
Me.cmd_liquidacion.Enabled = False

Screen.MousePointer = 11

If DGrid_pedidos.SelBookmarks.Count = 0 Then
    
    MsgBox "No se hallaron Pedidos marcados para Liquidar."
    Me.cmd_liquidacion.Enabled = True
    Screen.MousePointer = 0
    Exit Sub

End If

With resumen_mensual_ventas.Recordset
            
    mvBookMark = .Bookmark
    .MoveFirst
    .Find "fecha_venta =" & Me.txt_fecha.Text & ""
    
    If (.EOF) Then
     .MoveLast
     .AddNew
        !fecha_venta = Me.txt_fecha.Text
        !cant_10 = "0"
        !cant_18 = "0"
        !cant_27 = "0"
        !cant_43 = "0"
        !tot_10 = "0"
        !tot_18 = "0"
        !tot_27 = "0"
        !tot_43 = "0"
        
     .Update
    End If
End With

'Para cada registro seleccionado lo vamos a cancelar
'y generar su liquidación previa
For Each Var In DGrid_pedidos.SelBookmarks
    
    'Se crea la liquidación previa
    'en la tabla liquidacion se coloca todo
    'en estado vigente
    DGrid_pedidos.Bookmark = Var
    
    Me.hist_pedidos.Recordset.Bookmark = Var
    Me.liquidado.Refresh
    
    With liquidado.Recordset
        
        If Not (.BOF And .EOF) Then
        
            mvBookMark = .Bookmark
                
            .MoveLast
            
        End If
        
        'Añadimos la liquidacion del cliente
        .AddNew
        
        !id_pedido = DGrid_pedidos.Columns(2).Text
        '!id_pedido = pedidos.Recordset!id_pedido
        
        !usuario_liq = Usuario
        
        !codigo = DGrid_pedidos.Columns(3).Text
        'pedidos.Recordset!codigo
        
        'El estatus del pedido es cancelado
        !status = "CA"
        
        !fecha_liq = Date
        
        !monto = CCur(DGrid_pedidos.Columns(8).Text)
        '!monto = pedidos.Recordset!monto_fac
        
        .Update
    
    End With
    
            hist_pedidos.Recordset!status = "CA"
            hist_pedidos.Recordset!fecha_cancel = Date
            hist_pedidos.Recordset.Update
        
'aqui trabajo con el inventario
If txt_id_inst.Text = "10" Then
      
      With inventario.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "id_inst = 10"
           !cilindro = !cilindro - CInt(Me.txt_cant_pedido.Text)
        .Update
      End With
      Me.inventario.Refresh
      
      With materiales.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "codigo = 1"
           !cant_actual = !cant_actual - CInt(Me.txt_cant_pedido.Text)
        .Update
      End With
      Me.materiales.Refresh
   
   With res_inventario.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "codigo = 1"
           !cil_lleno = !cil_lleno - CInt(Me.txt_cant_pedido.Text)
           !cil_vacio = !cil_vacio + CInt(Me.txt_cant_pedido.Text)
         .Update
      End With
      Me.res_inventario.Refresh
     
   With resumen_mensual_ventas.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "fecha_venta =" & Me.txt_fecha.Text & ""
            !cant_10 = !cant_10 + CInt(Me.txt_cant.Text)
            !tot_10 = !tot_10 + CCur(Me.txt_monto.Text)
        .Update
   End With
      Me.resumen_mensual_ventas.Refresh
  End If
  
If txt_id_inst.Text = "18" Then
      
      With inventario.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "id_inst = 18"
           !cilindro = !cilindro - CInt(Me.txt_cant_pedido.Text)
        .Update
      End With
      Me.inventario.Refresh
      
      With materiales.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "codigo = 2"
           !cant_actual = !cant_actual - CInt(Me.txt_cant_pedido.Text)
        .Update
      End With
      Me.materiales.Refresh
   
      With res_inventario.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "codigo = 2"
           !cil_lleno = !cil_lleno - CInt(Me.txt_cant_pedido.Text)
           !cil_vacio = !cil_vacio + CInt(Me.txt_cant_pedido.Text)
         .Update
      End With
      Me.res_inventario.Refresh
      
   With resumen_mensual_ventas.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "fecha_venta =" & Me.txt_fecha.Text & ""
            !cant_18 = !cant_18 + CInt(Me.txt_cant.Text)
            !tot_18 = !tot_18 + CCur(Me.txt_monto.Text)
        .Update
   End With
      Me.resumen_mensual_ventas.Refresh
 End If
 
   If txt_id_inst.Text = "27" Then
      
      With inventario.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "id_inst = 27"
           !cilindro = !cilindro - CInt(Me.txt_cant_pedido.Text)
        .Update
      End With
      Me.inventario.Refresh
      
      With materiales.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "codigo = 3"
           !cant_actual = !cant_actual - CInt(Me.txt_cant_pedido.Text)
        .Update
      End With
      Me.materiales.Refresh
   
      With res_inventario.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "codigo = 3"
           !cil_lleno = !cil_lleno - CInt(Me.txt_cant_pedido.Text)
           !cil_vacio = !cil_vacio + CInt(Me.txt_cant_pedido.Text)
         .Update
      End With
      Me.res_inventario.Refresh
     
   With resumen_mensual_ventas.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "fecha_venta =" & Me.txt_fecha.Text & ""
            !cant_27 = !cant_27 + CInt(Me.txt_cant.Text)
            !tot_27 = !tot_27 + CCur(Me.txt_monto.Text)
        .Update
   End With
      Me.resumen_mensual_ventas.Refresh
 End If
   
   If txt_id_inst.Text = "43" Then
           
        With inventario.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "id_inst = 43"
           !cilindro = !cilindro - CInt(Me.txt_cant_pedido.Text)
        .Update
      End With
      Me.inventario.Refresh
      
      With materiales.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "codigo = 4"
           !cant_actual = !cant_actual - CInt(Me.txt_cant_pedido.Text)
        .Update
      End With
      Me.materiales.Refresh
   
      With res_inventario.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "codigo = 4"
           !cil_lleno = !cil_lleno - CInt(Me.txt_cant_pedido.Text)
           !cil_vacio = !cil_vacio + CInt(Me.txt_cant_pedido.Text)
         .Update
      End With
      Me.res_inventario.Refresh
   
   With resumen_mensual_ventas.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "fecha_venta =" & Me.txt_fecha.Text & ""
            !cant_43 = !cant_43 + CInt(Me.txt_cant.Text)
            !tot_43 = !tot_43 + CCur(Me.txt_monto.Text)
        .Update
   End With
      Me.resumen_mensual_ventas.Refresh
 End If
 
Next

'aqui actualizo la tabla tbl_resumen_mensual_ventas
      
'  With resumen_mensual_ventas.Recordset
'
'            mvBookMark = .Bookmark
'           .MoveFirst
'
''           .Find "fecha_venta =" & Me.txt_fecha.Text & ""
'
'           Do While ultimo = False
'
'           .Find "fecha_venta =" & Me.txt_fecha.Text & ""
'
'            If Not (.EOF) Then
'
'              If txt_inst.Text = "10" Then
'                  !fecha_venta = Me.txt_fecha.Text
'                    If (IsNull(!cant_10)) Then
'                    !cant_10 = CInt(Me.txt_cant.Text)
'                    !tot_10 = CCur(Me.txt_monto.Text)
'                    Else
'                      !cant_10 = !cant_10 + CInt(Me.txt_cant.Text)
'                      !tot_10 = !tot_10 + CCur(Me.txt_monto.Text)
'                  End If
'              End If
'
'              If txt_inst.Text = "18" Then
'                  !fecha_venta = Me.txt_fecha.Text
'                    If (IsNull(!cant_18)) Then
'                    !cant_18 = CInt(Me.txt_cant.Text)
'                    !tot_18 = CCur(Me.txt_monto.Text)
'                    Else
'                      !cant_18 = !cant_18 + CInt(Me.txt_cant.Text)
'                      !tot_18 = !tot_18 + CCur(Me.txt_monto.Text)
'                  End If
'              End If
'
'              If txt_inst.Text = "27" Then
'                  !fecha_venta = Me.txt_fecha.Text
'                    If (IsNull(!cant_27)) Then
'                    !cant_27 = CInt(Me.txt_cant.Text)
'                    !tot_27 = CCur(Me.txt_monto.Text)
'                    Else
'                      !cant_27 = !cant_27 + CInt(Me.txt_cant.Text)
'                      !tot_27 = !tot_27 + CCur(Me.txt_monto.Text)
'                  End If
'              End If
'
'              If txt_inst.Text = "43" Then
'                  !fecha_venta = Me.txt_fecha.Text
'                    If (IsNull(!cant_43)) Then
'                    !cant_43 = CInt(Me.txt_cant.Text)
'                    !tot_43 = CCur(Me.txt_monto.Text)
'                    Else
'                      !cant_43 = !cant_43 + CInt(Me.txt_cant.Text)
'                      !tot_43 = !tot_43 + CCur(Me.txt_monto.Text)
'                  End If
'              End If
'            .Update
' '        Else
'           If (.EOF) Then
'              ultimo = True
'           End If
'        End If
'   Loop
' End With

   
hist_pedidos.Refresh

Screen.MousePointer = 0


Exit Sub

control_error:
Screen.MousePointer = 0
    MsgBox Err.Description

End Sub

'With resumen_mensual_ventas.Recordset
'
'           mvBookMark = .Bookmark
'           .MoveFirst
'           .Find "fecha_venta =" & Me.txt_fecha.Text & ""
'
'             If Not (.EOF) Then
'
'              If txt_inst.Text = "10" Then
'                  !fecha_venta = Me.txt_fecha.Text
'                    If (IsNull(!cant_10)) Then
'                    !cant_10 = CInt(Me.txt_cant.Text)
'                    !tot_10 = CCur(Me.txt_monto.Text)
'                    Else
'                      !cant_10 = !cant_10 + CInt(Me.txt_cant.Text)
'                      !tot_10 = !tot_10 + CCur(Me.txt_monto.Text)
'                  End If
'              End If
'
'              If txt_inst.Text = "18" Then
'                  !fecha_venta = Me.txt_fecha.Text
'                    If (IsNull(!cant_18)) Then
'                    !cant_18 = CInt(Me.txt_cant.Text)
'                    !tot_18 = CCur(Me.txt_monto.Text)
'                    Else
'                      !cant_18 = !cant_18 + CInt(Me.txt_cant.Text)
'                      !tot_18 = !tot_18 + CCur(Me.txt_monto.Text)
'                  End If
'              End If
'
'              If txt_inst.Text = "27" Then
'                  !fecha_venta = Me.txt_fecha.Text
'                    If (IsNull(!cant_27)) Then
'                    !cant_27 = CInt(Me.txt_cant.Text)
'                    !tot_27 = CCur(Me.txt_monto.Text)
'                    Else
'                      !cant_27 = !cant_27 + CInt(Me.txt_cant.Text)
'                      !tot_27 = !tot_27 + CCur(Me.txt_monto.Text)
'                  End If
'              End If
'
'              If txt_inst.Text = "43" Then
'                  !fecha_venta = Me.txt_fecha.Text
'                    If (IsNull(!cant_43)) Then
'                    !cant_43 = CInt(Me.txt_cant.Text)
'                    !tot_43 = CCur(Me.txt_monto.Text)
'                    Else
'                      !cant_43 = !cant_43 + CInt(Me.txt_cant.Text)
'                      !tot_43 = !tot_43 + CCur(Me.txt_monto.Text)
'                  End If
'              End If
'            .Update
'      '  End If
'      'End With
'Else
'      If (.EOF) Then
'              mvBookMark = .Bookmark
'           .MoveLast
'
'        .AddNew
'          If txt_inst.Text = "10" Then
'               !fecha_venta = Date
'               !cant_10 = CInt(Me.txt_cant.Text)
'               !tot_10 = CCur(txt_monto.Text)
'          End If
'          If txt_inst.Text = "18" Then
'               !fecha_venta = Date
'               !cant_18 = CInt(Me.txt_cant.Text)
'               !tot_18 = CCur(txt_monto.Text)
'          End If
'          If txt_inst.Text = "27" Then
'               !fecha_venta = Date
'               !cant_27 = CInt(Me.txt_cant.Text)
'               !tot_27 = CCur(txt_monto.Text)
'          End If
'          If txt_inst.Text = "43" Then
'               !fecha_venta = Date
'               !cant_43 = CInt(Me.txt_cant.Text)
'               !tot_43 = CCur(txt_monto.Text)
'          End If
'        .Update
'       End If
'  '  End With
