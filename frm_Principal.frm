VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000E&
   Caption         =   "SISTEMA DE GESTIÓN Y CONTROL JER-GAS, C.A."
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   150
   ClientWidth     =   10170
   Icon            =   "frm_Principal.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "frm_Principal.frx":08CA
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   10170
      TabIndex        =   0
      Top             =   0
      Width           =   10170
      Begin VB.CommandButton cmd_salir 
         Height          =   580
         Left            =   6960
         Picture         =   "frm_Principal.frx":38AF
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salir del Sistema"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmd_ayudas 
         Height          =   580
         Left            =   6000
         Picture         =   "frm_Principal.frx":3DB1
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Ayuda"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmd_inventatio 
         Height          =   580
         Left            =   5040
         Picture         =   "frm_Principal.frx":42E5
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Control de Inventarios"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmd_estantes 
         Height          =   580
         Left            =   4080
         Picture         =   "frm_Principal.frx":4835
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Control de Estantes"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmd_clientes 
         Height          =   580
         Left            =   240
         Picture         =   "frm_Principal.frx":4D96
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Archivo de Clientes"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Cmd_pedidos 
         Height          =   580
         Left            =   2160
         Picture         =   "frm_Principal.frx":5366
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Control de Pedidos"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmd_choferes 
         Height          =   580
         Left            =   1200
         Picture         =   "frm_Principal.frx":57EA
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Archivo de Conductores"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmd_facturas 
         Height          =   580
         Left            =   3120
         Picture         =   "frm_Principal.frx":5D82
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Control de Liquidación"
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Menu Archivo 
      Caption         =   "&Archivos"
      Begin VB.Menu arch01 
         Caption         =   "Clientes"
         Shortcut        =   ^K
      End
      Begin VB.Menu arch011 
         Caption         =   "Clientes (Estante)"
         Shortcut        =   ^E
      End
      Begin VB.Menu linea1 
         Caption         =   "-"
      End
      Begin VB.Menu Arch02 
         Caption         =   "Conductores"
         Shortcut        =   ^O
      End
      Begin VB.Menu Arch03 
         Caption         =   "Rutas"
         Shortcut        =   ^R
      End
      Begin VB.Menu linea2 
         Caption         =   "-"
      End
      Begin VB.Menu Arch04 
         Caption         =   "Productos"
         Shortcut        =   ^D
      End
      Begin VB.Menu linea3 
         Caption         =   "-"
      End
      Begin VB.Menu Saliendo1 
         Caption         =   "Salir"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu movimiento 
      Caption         =   "&Movimientos"
      Begin VB.Menu Movi01 
         Caption         =   "Ingresar Pedidos"
         Shortcut        =   ^P
      End
      Begin VB.Menu movi03 
         Caption         =   "Ingresar Pedidos (Estantes)"
         Shortcut        =   ^I
      End
      Begin VB.Menu movi06 
         Caption         =   "Ingresar Ventas de Materiales"
         Shortcut        =   ^V
      End
      Begin VB.Menu linea21 
         Caption         =   "-"
      End
      Begin VB.Menu movi07 
         Caption         =   "Generar Facturas"
         Shortcut        =   ^F
      End
      Begin VB.Menu movi04 
         Caption         =   "Liquidar Pedidos"
         Shortcut        =   ^L
      End
      Begin VB.Menu linea22 
         Caption         =   "-"
      End
      Begin VB.Menu Movi02 
         Caption         =   "Nuevas Instalaciones"
         Shortcut        =   ^N
      End
   End
   Begin VB.Menu Reporte 
      Caption         =   "&Reportes"
      Begin VB.Menu Reportes 
         Caption         =   "Clientes"
         Begin VB.Menu Reporte01 
            Caption         =   "Reporte Maestro de Clientes (Código)"
         End
         Begin VB.Menu Reporte02 
            Caption         =   "Reporte Maestro de Clientes (Contratos de Instalación)"
         End
         Begin VB.Menu Reporte03 
            Caption         =   "Reporte Maestro de Clientes (Rutas)"
         End
         Begin VB.Menu Reporte04 
            Caption         =   "Reporte Maestro de Clientes (Tipo de Instalación)"
         End
         Begin VB.Menu Reporte05 
            Caption         =   "Reporte Maestro de Clientes (Relación de Últimos Pedidos)"
         End
         Begin VB.Menu clientes06 
            Caption         =   "Reporte Maestro de Clientes (Histórico de Pedidos)"
         End
      End
      Begin VB.Menu linea8 
         Caption         =   "-"
      End
      Begin VB.Menu Reporte06 
         Caption         =   "Facturas de Ventas Diarias"
      End
      Begin VB.Menu Reporte08 
         Caption         =   "Resumen de Pedidos (Rutas)"
      End
      Begin VB.Menu Reporte07 
         Caption         =   "Resumen de Ventas Diarias"
      End
      Begin VB.Menu reporte20 
         Caption         =   "Relación de Ventas Diarias"
      End
      Begin VB.Menu linea4 
         Caption         =   "-"
      End
      Begin VB.Menu Reporte09 
         Caption         =   "Resumen Mensual de Ventas (Rutas) "
      End
      Begin VB.Menu reporte10 
         Caption         =   "Resumen Mensual de Ventas (Cilindros)"
      End
      Begin VB.Menu reporte16 
         Caption         =   "Resumen Mensual de Ventas (Materiales)"
      End
      Begin VB.Menu Reporte17 
         Caption         =   "Resumen Mensual de Ventas (Nuevas Instalaciones)"
      End
      Begin VB.Menu linea5 
         Caption         =   "-"
      End
      Begin VB.Menu Control1 
         Caption         =   "Control de Estantes"
         Begin VB.Menu Reporte11 
            Caption         =   "Resumen de Contratos"
            Enabled         =   0   'False
         End
         Begin VB.Menu Reporte12 
            Caption         =   "Resumen de Ventas (10 kgs)"
         End
         Begin VB.Menu Reporte13 
            Caption         =   "Históricos de Pedidos"
         End
      End
      Begin VB.Menu Control2 
         Caption         =   "Control de Inventarios"
         Begin VB.Menu reporte14 
            Caption         =   "Resumen Mensual de Rotación de Producto"
         End
         Begin VB.Menu Reporte15 
            Caption         =   "Resumen Mensual de Inventario"
         End
      End
   End
   Begin VB.Menu Herramienta 
      Caption         =   "&Herramientas"
      Begin VB.Menu conf_despacho 
         Caption         =   "Configurar Despachos"
         Shortcut        =   ^C
      End
      Begin VB.Menu liberar 
         Caption         =   "Liberar Factura"
         Shortcut        =   ^U
      End
      Begin VB.Menu exportar_pedidos 
         Caption         =   "Exportar a Excel los Pedidos"
      End
      Begin VB.Menu linea12 
         Caption         =   "-"
      End
      Begin VB.Menu respaldo 
         Caption         =   "Respaldo Mensual"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu Ayuda 
      Caption         =   "A&yuda"
      Begin VB.Menu Ayuda01 
         Caption         =   "Contenido"
         Shortcut        =   ^H
      End
      Begin VB.Menu linea6 
         Caption         =   "-"
      End
      Begin VB.Menu Ayuda03 
         Caption         =   "Acerca de"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Arch01_Click()
frm_clientes.Show
End Sub

Private Sub Arch011_Click()
frm_clientes_est.Show
End Sub

Private Sub Arch02_Click()
frm_choferes.Show
End Sub

Private Sub Arch03_Click()
frm_rutas.Show
End Sub

Private Sub Arch04_Click()
frm_materiales.Show
End Sub

Private Sub cmd_ayudas_Click()
frm_menu_ayuda.Show
End Sub

Private Sub cmd_choferes_Click()
frm_choferes.Show
End Sub

Private Sub cmd_clientes_Click()
frm_clientes.Show
End Sub

Private Sub cmd_estantes_Click()
frm_clientes_est.Show
End Sub

Private Sub cmd_facturas_Click()
frm_liquidar_pedidos.Show
End Sub

Private Sub cmd_inventatio_Click()
frm_materiales.Show
End Sub

Private Sub cmd_pedidos_Click()
frm_pedidos.Show
End Sub

Private Sub cmd_salir_Click()
Unload Me
End Sub

Private Sub conf_despacho_Click()
frm_config_despacho.Show
End Sub

Private Sub exportar_pedidos_Click()
frm_exportar_excel.Show
End Sub

Private Sub liberar_Click()
frm_liberar_planilla.Show
End Sub

Private Sub Movi01_Click()
frm_pedidos.Show
End Sub

Private Sub Movi02_Click()
frm_instalaciones.Show
End Sub

Private Sub Movi03_Click()
frm_pedidos_estantes.Show
End Sub

Private Sub movi04_Click()
frm_liquidar_pedidos.Show
End Sub

Private Sub movi06_Click()
frm_ventas_materiales2.Show
End Sub

Private Sub movi07_Click()
frm_genera_factura.Show
End Sub

Private Sub respaldo_Click()
frm_respaldo_mensual.Show
End Sub

Private Sub Saliendo1_Click()
Unload Me
End Sub

Private Sub salir_Click()
Unload Me
End Sub

'LOS REPORTES

Private Sub Reporte01_Click()
rpt_clientes1.Show
End Sub

Private Sub Reporte02_Click()
rpt_clientes2A.Show
End Sub

Private Sub Reporte03_Click()
rpt_clientes3.Show
End Sub

Private Sub Reporte04_Click()
rpt_clientes4.Show
End Sub

Private Sub Reporte05_Click()
rpt_clientes5.Show
End Sub

Private Sub clientes06_Click()
frm_historico_pedidos_clientes.Show
End Sub

Private Sub Reporte06_Click()
rpt_factura2.Show
End Sub

Private Sub Reporte07_Click()
frm_resumen_ventas.Show
End Sub

Private Sub Reporte08_Click()
frm_resumen_pedidos.Show
End Sub

Private Sub Reporte09_Click()
frm_resumen_mensual_rutas.Show
End Sub

Private Sub Reporte10_Click()
frm_resumen_mensual_ventas.Show
End Sub

Private Sub Reporte12_Click()
frm_resumen_pedidos_estantes.Show
End Sub

Private Sub Reporte13_Click()
frm_historico_pedidos_estantes.Show
End Sub

Private Sub Reporte14_Click()
rpt_inventario.Show
'frm_resumen_mensual_inventario1.Show
End Sub

Private Sub Reporte15_Click()
rpt_resumen_inventario.Show
End Sub

Private Sub reporte16_Click()
frm_resumen_mensual_materiales.Show
End Sub

Private Sub reporte17_Click()
frm_resumen_mensual_instalaciones.Show
End Sub

Private Sub reporte20_Click()
frm_relacion_ventas.Show
End Sub

'LAS AYUDAS

Private Sub Ayuda01_Click()
frm_menu_ayuda.Show
End Sub

Private Sub Ayuda03_Click()
frm_acerca_de.Show
End Sub

