VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_exportar_excel 
   Caption         =   "Exportar Pedidos"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   6015
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   8415
      Begin VB.CommandButton cmd_crear 
         Caption         =   "Crear Archivo"
         Height          =   495
         Left            =   2160
         TabIndex        =   18
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CommandButton cmd_salir 
         Caption         =   "Salir"
         Height          =   495
         Left            =   2040
         TabIndex        =   16
         Top             =   4680
         Width           =   1815
      End
      Begin VB.CommandButton cmd_seleccione 
         Caption         =   "Seleccione el Archivo"
         Height          =   495
         Left            =   2040
         TabIndex        =   15
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton cmd_estado 
         Caption         =   "Estado de Cuenta"
         Enabled         =   0   'False
         Height          =   735
         Left            =   4920
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   6240
         Width           =   1455
      End
      Begin VB.CommandButton cmd_clientes 
         Caption         =   "Editar Clientes"
         Enabled         =   0   'False
         Height          =   735
         Left            =   6360
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   6240
         Width           =   1455
      End
      Begin VB.CommandButton cmd_liquidacion 
         Caption         =   "Liquidación"
         Enabled         =   0   'False
         Height          =   735
         Left            =   600
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   6240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "Cerrar"
         Height          =   735
         Left            =   7800
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   6240
         Width           =   1455
      End
      Begin VB.CommandButton cmd_eliminar 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Height          =   735
         Left            =   3480
         TabIndex        =   8
         Top             =   6240
         Width           =   1455
      End
      Begin VB.TextBox txt_monto 
         Height          =   375
         Left            =   2640
         TabIndex        =   7
         Top             =   6840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txt_fecha 
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   6840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txt_cant 
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   6840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txt_inst 
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   6840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text1 
         DataField       =   "fecha_venta"
         DataSource      =   "resumen_mensual_ventas"
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   6360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmd_ventas_materiales 
         Caption         =   "Ventas de Materiales"
         Enabled         =   0   'False
         Height          =   735
         Left            =   2040
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   6240
         Width           =   1455
      End
      Begin MSComDlg.CommonDialog cdlBox 
         Left            =   6000
         Top             =   4200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "(Excel *.xls)|*.xls"
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   720
         TabIndex        =   13
         Top             =   5400
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Max             =   5000
      End
      Begin MSComCtl2.DTPicker fecha_desde 
         Height          =   375
         Left            =   2280
         TabIndex        =   21
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   17170433
         CurrentDate     =   37704
      End
      Begin MSComCtl2.DTPicker fecha_hasta 
         Height          =   375
         Left            =   5520
         TabIndex        =   22
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   17170433
         CurrentDate     =   37704
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha Hasta:"
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
         Left            =   4080
         TabIndex        =   24
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha desde:"
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
         Left            =   840
         TabIndex        =   23
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Lbl_archivo 
         Height          =   375
         Left            =   840
         TabIndex        =   20
         Top             =   1440
         Width           =   7575
      End
      Begin VB.Label Label3 
         Caption         =   "Para ver el archivo búsquelo y ábralo con Microsoft Excel."
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
         Left            =   840
         TabIndex        =   19
         Top             =   4200
         Width           =   7335
      End
      Begin VB.Label Label2 
         Caption         =   "El siguiente botón procesa los datos:"
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
         Left            =   840
         TabIndex        =   17
         Top             =   1800
         Width           =   7335
      End
      Begin VB.Label Label1 
         Caption         =   "Indique el archivo excel (en blanco) en donde se copiara los pedidos."
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
         Left            =   840
         TabIndex        =   14
         Top             =   480
         Width           =   7335
      End
   End
   Begin MSAdodcLib.Adodc resumen_pedidos 
      Height          =   495
      Left            =   2160
      Top             =   840
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      RecordSource    =   "resumen_pedidos"
      Caption         =   "resumen_pedidos"
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
      Caption         =   "  EXPORTAR PEDIDOS"
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
      Top             =   480
      Width           =   7695
   End
End
Attribute VB_Name = "frm_exportar_excel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, fso, strimagen, cont, colum, valor, cel
Dim EXL As Excel.Application

Dim W As Excel.Workbook

Private Sub cmd_crear_Click()
    Dim S As Excel.Worksheet
    Set S = W.Sheets("Hoja1")
    cont = 1
    'cmd_cargar.Enabled = False
    ProgressBar1.Visible = True
    ProgressBar1.Value = cont
    colum = "A" & cont & ""
    S.Range(colum).Value = "cliente"
    S.Range("B1").Value = "direccion"
    S.Range("C1").Value = "id_ruta"
    S.Range("D1").Value = "id_pedido"
    S.Range("E1").Value = "codigo"
    S.Range("F1").Value = "Status"
    S.Range("G1").Value = "fecha_pedido"
    S.Range("H1").Value = "id_inst"
    S.Range("I1").Value = "cant_pedido"
    S.Range("J1").Value = "monto_fac"
    S.Range("K1").Value = "nombre"
    S.Range("L1").Value = "fecha_cancel"
    resumen_pedidos.CommandType = adCmdText
    resumen_pedidos.RecordSource = "select * from resumen_pedidos WHERE fecha_cancel >='" & fecha_desde & "' AND fecha_cancel <='" & fecha_hasta & "'"
    'maxcont = resumen_pedidos.Recordset.
    resumen_pedidos.Refresh
    Me.ProgressBar1.Max = maxcont
    
    While Not resumen_pedidos.Recordset.EOF
    cont = cont + 1
            S.Range("A" & cont & "").Value = resumen_pedidos.Recordset!cliente
            S.Range("B" & cont & "").Value = resumen_pedidos.Recordset!direccion
            S.Range("C" & cont & "").Value = resumen_pedidos.Recordset!id_ruta
            S.Range("D" & cont & "").Value = resumen_pedidos.Recordset!id_pedido
            S.Range("E" & cont & "").Value = resumen_pedidos.Recordset!codigo
            S.Range("F" & cont & "").Value = resumen_pedidos.Recordset!Status
            S.Range("G" & cont & "").Value = resumen_pedidos.Recordset!fecha_pedido
            S.Range("H" & cont & "").Value = resumen_pedidos.Recordset!id_inst
            S.Range("I" & cont & "").Value = resumen_pedidos.Recordset!cant_pedido
            S.Range("J" & cont & "").Value = resumen_pedidos.Recordset!monto_fac
            S.Range("K" & cont & "").Value = resumen_pedidos.Recordset!Nombre
            S.Range("L" & cont & "").Value = resumen_pedidos.Recordset!fecha_cancel
        resumen_pedidos.Recordset.MoveNext
        ProgressBar1.Value = I
    Wend
    
    Set S = Nothing
    W.Save
    W.Close
    Set W = Nothing
    Set EXL = Nothing
End Sub

Private Sub cmd_salir_Click()
Unload Me
End Sub

Private Sub cmd_seleccione_Click()
Set EXL = New Excel.Application
cdlBox.ShowOpen
Set fso = CreateObject("Scripting.FileSystemObject")
If cdlBox.FileName <> "" Then
    Set a = fso.GetFile(cdlBox.FileName)
    Set W = EXL.Workbooks.Open(a)
    Lbl_archivo.Caption = a
End If
End Sub

Private Sub Form_Load()
fecha_desde.Value = Date
fecha_hasta.Value = Date
End Sub

Private Sub Form_Resize()

'Call Mover_der(Me, Label_titulo, 0)
'Call Mover_centrado(Me, Frame1)

End Sub

