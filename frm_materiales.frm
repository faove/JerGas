VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_materiales 
   Caption         =   "SISTEMA DE GESTIÓN Y CONTROL JER-GAS, C.A."
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9465
   Icon            =   "frm_materiales.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   9465
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text3 
      DataField       =   "codigo"
      DataSource      =   "resumen_inv1"
      Height          =   285
      Left            =   5760
      TabIndex        =   50
      Text            =   "Text3"
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text2 
      DataField       =   "codigo"
      DataSource      =   "resumen_inv2"
      Height          =   285
      Left            =   5760
      TabIndex        =   49
      Text            =   "Text2"
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame12 
      Height          =   1215
      Left            =   11760
      TabIndex        =   47
      Top             =   8160
      Width           =   1695
      Begin VB.CommandButton cmd_actualiza_precio 
         Caption         =   "&Actualizar Precios de Venta"
         Height          =   855
         Left            =   120
         TabIndex        =   48
         Top             =   200
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc ingresos 
      Height          =   330
      Left            =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
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
      RecordSource    =   "tbl_ingresos_materiales"
      Caption         =   "ingresos"
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
   Begin VB.TextBox txt_id_inst 
      DataField       =   "id_inst"
      DataSource      =   "materiales"
      Height          =   375
      Left            =   2160
      TabIndex        =   46
      Text            =   "Text3"
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txt_contador 
      DataField       =   "codigo"
      DataSource      =   "ingresos"
      Height          =   285
      Left            =   2160
      TabIndex        =   45
      Text            =   "Text2"
      Top             =   690
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      DataField       =   "id_inst"
      DataSource      =   "inventario"
      Height          =   375
      Left            =   2160
      TabIndex        =   44
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc inventario 
      Height          =   330
      Left            =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
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
   Begin VB.Frame Frame8 
      Caption         =   "INGRESAR MATERIALES"
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
      Height          =   2055
      Left            =   9360
      TabIndex        =   38
      Top             =   1920
      Width           =   4095
      Begin VB.CommandButton cmdguardar2 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   660
         Left            =   2040
         TabIndex        =   42
         ToolTipText     =   "Guardar Materiales en Stock Actual"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdingresar 
         Caption         =   "&Ingresar"
         Height          =   660
         Left            =   840
         TabIndex        =   41
         ToolTipText     =   "Ingresa Materiales al Stock Actual"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Frame Frame11 
         Caption         =   "Cantidad"
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
         Height          =   765
         Left            =   2400
         TabIndex        =   40
         ToolTipText     =   "Cantidad"
         Top             =   360
         Width           =   1335
         Begin VB.TextBox txt_ingresa 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   380
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Fecha de Ingreso"
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
         Height          =   765
         Left            =   240
         TabIndex        =   39
         ToolTipText     =   "Fecha de Ingreso"
         Top             =   360
         Width           =   1815
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   59506689
            CurrentDate     =   39394
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Código"
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
      Height          =   645
      Left            =   600
      TabIndex        =   20
      ToolTipText     =   "Código de Material"
      Top             =   2280
      Width           =   1215
      Begin VB.TextBox txt_codigo 
         Alignment       =   2  'Center
         DataField       =   "codigo"
         DataSource      =   "materiales"
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
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "INVENTARIO (STOCK ACTUAL )"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   5205
      Left            =   360
      TabIndex        =   25
      ToolTipText     =   "Inventario (stock Actual)"
      Top             =   4200
      Width           =   11175
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frm_materiales.frx":08CA
         Height          =   3975
         Left            =   120
         TabIndex        =   28
         Top             =   840
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   7011
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   -2147483633
         BorderStyle     =   0
         Enabled         =   0   'False
         ColumnHeaders   =   0   'False
         HeadLines       =   1
         RowHeight       =   19
         TabAction       =   1
         RowDividerStyle =   0
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "codigo"
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "descripcion"
            Caption         =   ""
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
            DataField       =   "cant_inicial"
            Caption         =   ""
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
            DataField       =   "cant_actual"
            Caption         =   ""
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
            DataField       =   "precio"
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   " #.##0,00;( #.##0,00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "cant_minima"
            Caption         =   ""
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
            DataField       =   "cant_maxima"
            Caption         =   ""
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
            DataField       =   "fecha_pedido"
            Caption         =   ""
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
            DataField       =   "kit"
            Caption         =   "kit"
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
            MarqueeStyle    =   0
            ScrollBars      =   0
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            Size            =   2
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   705,26
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2775,118
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   1200,189
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   1200,189
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1200,189
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   854,929
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               ColumnWidth     =   854,929
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               ColumnWidth     =   1395,213
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
               ColumnWidth     =   2775,118
            EndProperty
         EndProperty
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "FECHA INGRESO"
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
         Height          =   375
         Left            =   9120
         TabIndex        =   36
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "MAX"
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
         Height          =   255
         Left            =   8400
         TabIndex        =   35
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "MIN"
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
         Height          =   255
         Left            =   7440
         TabIndex        =   34
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "PRECIO"
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
         Height          =   255
         Left            =   6360
         TabIndex        =   33
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "C/ACTUAL"
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
         Height          =   255
         Left            =   5160
         TabIndex        =   32
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "C/INICIAL"
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
         Height          =   255
         Left            =   3960
         TabIndex        =   31
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "DESCRIPCIÓN"
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
         Height          =   255
         Left            =   1680
         TabIndex        =   30
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "CÓDIGO"
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
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "Cant. Inicial"
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
      Left            =   3120
      TabIndex        =   23
      ToolTipText     =   "Cantidad Inicial"
      Top             =   3120
      Width           =   1335
      Begin VB.TextBox txt_cant_inicial 
         Alignment       =   2  'Center
         DataField       =   "cant_inicial"
         DataSource      =   "materiales"
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
         Width           =   1095
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Precio de Instalación"
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
      Left            =   600
      TabIndex        =   22
      ToolTipText     =   "Precio de Instalación"
      Top             =   3120
      Width           =   2295
      Begin VB.TextBox txt_precio 
         Alignment       =   2  'Center
         DataField       =   "precio"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "materiales"
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
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Descripcion"
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
      Left            =   2040
      TabIndex        =   21
      ToolTipText     =   "Descripción"
      Top             =   2280
      Width           =   3855
      Begin VB.TextBox txt_descripcion 
         DataField       =   "descripcion"
         DataSource      =   "materiales"
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
         Width           =   3615
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Cant. Min"
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
      Left            =   6120
      TabIndex        =   19
      ToolTipText     =   "Cantidad Mínima "
      Top             =   2280
      Width           =   1335
      Begin VB.TextBox txt_cant_min 
         Alignment       =   2  'Center
         DataField       =   "cant_minima"
         DataSource      =   "materiales"
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
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>|"
      Height          =   615
      Index           =   3
      Left            =   8280
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      Height          =   615
      Index           =   2
      Left            =   7680
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<<"
      Height          =   615
      Index           =   1
      Left            =   7080
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "|<<"
      Height          =   615
      Index           =   0
      Left            =   6480
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3120
      Width           =   615
   End
   Begin VB.Frame Frame17 
      Height          =   3975
      Left            =   11760
      TabIndex        =   8
      Top             =   4200
      Width           =   1695
      Begin VB.CommandButton cmdsalir 
         Caption         =   "C&errar"
         Height          =   660
         Left            =   240
         TabIndex        =   14
         ToolTipText     =   "Cerrar y Volver al Menú Principal"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton cmdeliminar 
         Caption         =   "&Eliminar"
         Height          =   660
         Left            =   240
         TabIndex        =   13
         ToolTipText     =   "Elimina de la Base de Datos a un Material"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdmodificar 
         Caption         =   "&Modificar"
         Height          =   660
         Left            =   240
         TabIndex        =   12
         ToolTipText     =   "Cambiar Característica de un Material"
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdguardar 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   660
         Left            =   240
         TabIndex        =   11
         ToolTipText     =   "Para Salvar el Material Agregado o Modificado"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdcancelar 
         Caption         =   "&Cancelar"
         Height          =   660
         Left            =   240
         TabIndex        =   10
         ToolTipText     =   "Pulse este botón si desea Cancelar el Material Agregado"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdagregar 
         Caption         =   "&Agregar"
         Height          =   660
         Left            =   240
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Botón para Agregar un Nuevo Usuario"
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc materiales 
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
   Begin VB.Frame Frame4 
      Caption         =   "DESCRIPCIÓN DE MATERIALES"
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
      Height          =   2085
      Left            =   360
      TabIndex        =   26
      ToolTipText     =   "Suministre el Código para su ingreso."
      Top             =   1920
      Width           =   8895
      Begin VB.Frame Frame7 
         Caption         =   "Cant. Actual"
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
         Left            =   4320
         TabIndex        =   37
         ToolTipText     =   "Cantidad Actual"
         Top             =   1200
         Width           =   1335
         Begin VB.TextBox txt_cant_actual 
            Alignment       =   2  'Center
            DataField       =   "cant_actual"
            DataSource      =   "materiales"
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
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Cant. Max"
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
         Left            =   7320
         TabIndex        =   27
         ToolTipText     =   "Cantidad Máxima"
         Top             =   360
         Width           =   1335
         Begin VB.TextBox txt_cant_max 
            Alignment       =   2  'Center
            DataField       =   "cant_maxima"
            DataSource      =   "materiales"
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
            Width           =   1095
         End
      End
   End
   Begin MSAdodcLib.Adodc resumen_inv2 
      Height          =   330
      Left            =   3240
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
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
      RecordSource    =   "tbl_resumen_inventario2"
      Caption         =   "resumen_inv2"
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
   Begin MSAdodcLib.Adodc resumen_inv1 
      Height          =   330
      Left            =   3240
      Top             =   360
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
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
      Caption         =   "resumen_inv1"
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
      Caption         =   "  CONTROL DE MATERIALES"
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
      TabIndex        =   24
      Top             =   240
      Width           =   8655
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000003&
      BorderColor     =   &H80000003&
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   0
      Top             =   960
      Width           =   15615
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
      Width           =   15345
   End
End
Attribute VB_Name = "frm_materiales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_actualiza_precio_Click()
frm_actualiza_precio.Show
End Sub

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
    txt_descripcion.Locked = False
    txt_cant_inicial.Locked = False
    txt_cant_actual.Locked = False
    txt_cant_min.Locked = False
    txt_cant_max.Locked = False
    txt_precio.Locked = False
    
    With materiales.Recordset
    
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
    
    materiales.Recordset.CancelUpdate
            
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
    txt_descripcion.Locked = True
    txt_cant_inicial.Locked = True
    txt_cant_actual.Locked = True
    txt_cant_min.Locked = True
    txt_cant_max.Locked = True
    txt_precio.Locked = True

    Exit Sub    ' Salir para evitar el controlador.

ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "Control del Materialeses")
        
    End Select

 End Sub

Private Sub cmdeliminar_Click()
  'esto puede producir un error si elimina el último
  'registro o el único registro del recordset

On Error GoTo ControlError
    respuesta = MsgBox("¿Desea Eliminar el Registro?", vbYesNo)
    If respuesta = vbYes Then
        materiales.Recordset.Delete
        materiales.Recordset.MoveNext
    End If

    Exit Sub    ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "Control de Materialeses")
    End Select
End Sub

Private Sub cmdingresar_Click()
Me.Frame10.Enabled = True
Me.Frame11.Enabled = True
cmdguardar2.Enabled = True
Me.txt_ingresa.Enabled = True
Me.DTPicker1.Enabled = True
Me.cmdingresar.Enabled = False
txt_ingresa.Locked = False
txt_ingresa.Text = ""
End Sub

Private Sub cmdguardar2_Click()
Dim Total As Integer
Dim contador As Integer


With materiales.Recordset

        mvBookMark = .Bookmark
       
       Total = CInt(txt_cant_actual) + CInt(txt_ingresa.Text)
       Me.txt_cant_actual = Total
       !fecha_pedido = Me.DTPicker1.Value
        .Update

        .Bookmark = mvBookMark

   If txt_codigo.Text >= 1 And txt_codigo.Text <= 4 Then
         
         With resumen_inv1.Recordset
           mvBookMark = .Bookmark
           .MoveFirst
           .Find "codigo =" & Me.txt_codigo.Text & ""
              !cil_lleno = !cil_lleno + CInt(Me.txt_ingresa.Text)
              !cil_vacio = !cil_vacio - CInt(Me.txt_ingresa.Text)
           .Update
         End With
           Me.resumen_inv1.Refresh
     End If
   
   If txt_codigo.Text >= 5 And txt_codigo.Text <= 14 Then
         
         With resumen_inv2.Recordset
           mvBookMark = .Bookmark
           .MoveFirst
           .Find "codigo =" & Me.txt_codigo.Text & ""
              !cant_actual = CInt(Me.txt_cant_actual.Text)
           .Update
         End With
           Me.resumen_inv2.Refresh
     End If

   If txt_codigo.Text = "1" Then
      With inventario.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "id_inst = 10"
           !cilindro = CInt(Me.txt_cant_actual.Text)
        .Update
       End With
   End If

   If txt_codigo.Text = "2" Then
      With inventario.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "id_inst = 18"
           !cilindro = CInt(Me.txt_cant_actual.Text)
        .Update
      End With
   End If

   If txt_codigo.Text = "3" Then
      With inventario.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "id_inst = 27"
           !cilindro = CInt(Me.txt_cant_actual.Text)
        .Update
      End With
   End If

    If txt_codigo.Text = "4" Then
      With inventario.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "id_inst = 43"
           !cilindro = CInt(Me.txt_cant_actual.Text)
        .Update
      End With
   End If

   If txt_codigo.Text = "5" Then
      With inventario.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "id_inst = 10"
           !regulador = CInt(Me.txt_cant_actual.Text)
        .Update
      End With
   End If

   If txt_codigo.Text = "6" Then
      With inventario.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "id_inst = 18"
           !regulador = CInt(Me.txt_cant_actual.Text)
        .Update
       End With
   End If

   If txt_codigo.Text = "7" Then
      With inventario.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "id_inst = 27"
           !regulador = CInt(Me.txt_cant_actual.Text)
        .Update
      End With
   End If

   If txt_codigo.Text = "8" Then
      With inventario.Recordset
        mvBookMark = .Bookmark
        .MoveFirst
        .Find "id_inst = 43"
           !regulador = CInt(Me.txt_cant_actual.Text)
        .Update
      End With
   End If

      
   With ingresos.Recordset
        mvBookMark = .Bookmark
        .MoveLast
        .AddNew
           !codigo = Me.txt_codigo.Text
           !fecha_ingreso = Me.DTPicker1.Value
           !descripcion = Me.txt_descripcion.Text
           !cant_ingreso = Me.txt_ingresa.Text
           !cant_fecha = Me.txt_cant_actual.Text
           !id_inst = Me.txt_id_inst.Text
                 
         .Update
      End With

End With

Me.Frame10.Enabled = False
Me.Frame11.Enabled = False
cmdguardar2.Enabled = False
Me.txt_ingresa.Enabled = False
Me.DTPicker1.Enabled = False
Me.cmdingresar.Enabled = True
txt_ingresa.Locked = True

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
    txt_descripcion.Locked = False
    txt_cant_inicial.Locked = False
    txt_cant_actual.Locked = False
    txt_cant_min.Locked = False
    txt_cant_max.Locked = False
    txt_precio.Locked = False

End Sub

Private Sub cmdguardar_Click()
Dim fec As Date
Dim ano As Date
Dim strquery As String
Dim bandera As Boolean
Dim abc, ncliente, contrato

On Error GoTo ControlError
    
    
    If IsNull(txt_descripcion.Text) Or txt_descripcion.Text = "" Then
    
        MsgBox "Descripción del Material no puede ser nulo, por favor verifique ", vbInformation, "JerGas"
             
        Me.txt_descripcion.SetFocus
             
        Exit Sub
    End If
    
    If IsNull(Me.txt_cant_inicial.Text) Or txt_cant_inicial.Text = "" Then
    
        MsgBox "Debe Existir una Cantidad Inicial, por favor verifique ", vbInformation, "JerGas"
             
        Me.txt_cant_inicial.SetFocus
             
        Exit Sub
    End If
    
    If IsNull(Me.txt_precio.Text) Or txt_precio.Text = "" Then
    
        MsgBox "Debe Colocar el Precio del Material, por favor verifique ", vbInformation, "JerGas"
             
        Me.txt_precio.SetFocus
             
        Exit Sub
    End If
    
    With materiales.Recordset

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
            v = MsgBox("Formato No Válido", vbOKOnly, "Control de Materiales")
        Case 3314
            MsgBox "Verifique", vbOKOnly, "Control de Materiales"
        Case 524
            MsgBox "Verifique la Descripción", vbOKOnly, "Control de Materiales"
        Case -2147467259
            MsgBox "Error, Código suministrado ya existe", vbOKOnly, "Control de Materiales"
        Case -2147217842
            MsgBox "Error, cancele la operación y vuelva a intentarlo", vbOKOnly, "Control de Materiales"
        Case -2147217887
            MsgBox "Error, al crear histórico, se recomienda borrar el registro y volverlo a crear", vbOKOnly, "Control de Materiales"
    End Select
End Sub

Private Sub Form_Load()
    cmdcancelar.Visible = False
Me.Command1(0).Enabled = True
Me.Command1(1).Enabled = True
Me.Command1(2).Enabled = True
Me.Command1(3).Enabled = True
Me.Frame10.Enabled = False
Me.Frame11.Enabled = False
Me.txt_ingresa.Enabled = False
Me.DTPicker1.Enabled = False

txt_codigo.Locked = True
txt_cant_inicial.Locked = True
txt_cant_actual.Locked = True
txt_descripcion.Locked = True
txt_cant_min.Locked = True
txt_cant_max.Locked = True
txt_precio.Locked = True
End Sub

Private Sub cmdsalir_Click()
  Unload Me
End Sub

Private Sub Command1_Click(Index As Integer)
On Error GoTo ControlError

Select Case Index
    Case 0
       materiales.Recordset.MoveFirst
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
         Me.cmd_actualiza_precio.FontBold = False
    Case 1
       materiales.Recordset.MovePrevious
       Command1(2).Enabled = True
       Command1(3).Enabled = True
         Me.cmdsalir.FontBold = False
         Me.cmdguardar.FontBold = False
         Me.cmdeliminar.FontBold = False
         Me.cmdcancelar.FontBold = False
         Me.cmdmodificar.FontBold = False
         Me.cmdagregar.FontBold = False
         Me.cmd_actualiza_precio.FontBold = False
       
       If materiales.Recordset.BOF = True Then
        materiales.Recordset.MoveFirst
        Command1(0).Enabled = False
        Command1(1).Enabled = False
         Me.cmdsalir.FontBold = False
         Me.cmdguardar.FontBold = False
         Me.cmdeliminar.FontBold = False
         Me.cmdcancelar.FontBold = False
         Me.cmdmodificar.FontBold = False
         Me.cmdagregar.FontBold = False
         Me.cmd_actualiza_precio.FontBold = False
      Else
        End If
    Case 2
       materiales.Recordset.MoveNext
       Command1(0).Enabled = True
       Command1(1).Enabled = True
         Me.cmdsalir.FontBold = False
         Me.cmdguardar.FontBold = False
         Me.cmdeliminar.FontBold = False
         Me.cmdcancelar.FontBold = False
         Me.cmdmodificar.FontBold = False
         Me.cmdagregar.FontBold = False
         Me.cmd_actualiza_precio.FontBold = False
       
       If materiales.Recordset.EOF = True Then
         Command1(2).Enabled = False
         Command1(3).Enabled = False
         Me.cmdsalir.FontBold = False
         Me.cmdguardar.FontBold = False
         Me.cmdeliminar.FontBold = False
         Me.cmdcancelar.FontBold = False
         Me.cmdmodificar.FontBold = False
         Me.cmdagregar.FontBold = False
         Me.cmd_actualiza_precio.FontBold = False

         materiales.Recordset.MoveLast
       Else
       End If
    Case 3
       materiales.Recordset.MoveLast
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
         Me.cmd_actualiza_precio.FontBold = False

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
Me.cmdagregar.FontBold = False
Me.cmdguardar.FontBold = False
Me.cmdeliminar.FontBold = False
Me.cmdcancelar.FontBold = False
Me.cmdsalir.FontBold = False
Me.cmd_actualiza_precio.FontBold = False
End Sub

Private Sub cmdguardar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdguardar.FontBold = True
Me.cmdagregar.FontBold = False
Me.cmdmodificar.FontBold = False
Me.cmdeliminar.FontBold = False
Me.cmdcancelar.FontBold = False
Me.cmdsalir.FontBold = False
Me.cmd_actualiza_precio.FontBold = False
End Sub

Private Sub cmdeliminar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdeliminar.FontBold = True
Me.cmdagregar.FontBold = False
Me.cmdguardar.FontBold = False
Me.cmdmodificar.FontBold = False
Me.cmdcancelar.FontBold = False
Me.cmdsalir.FontBold = False
Me.cmd_actualiza_precio.FontBold = False
End Sub

Private Sub cmdcancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdcancelar.FontBold = True
Me.cmdagregar.FontBold = False
Me.cmdguardar.FontBold = False
Me.cmdeliminar.FontBold = False
Me.cmdmodificar.FontBold = False
Me.cmdsalir.FontBold = False
Me.cmd_actualiza_precio.FontBold = False
End Sub

Private Sub cmdagregar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdagregar.FontBold = True
Me.cmdguardar.FontBold = False
Me.cmdeliminar.FontBold = False
Me.cmdcancelar.FontBold = False
Me.cmdmodificar.FontBold = False
Me.cmdsalir.FontBold = False
Me.cmd_actualiza_precio.FontBold = False
End Sub

Private Sub cmdsalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdsalir.FontBold = True
Me.cmdguardar.FontBold = False
Me.cmdeliminar.FontBold = False
Me.cmdcancelar.FontBold = False
Me.cmdmodificar.FontBold = False
Me.cmdagregar.FontBold = False
Me.cmd_actualiza_precio.FontBold = False
End Sub

Private Sub cmd_actualiza_precio_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdsalir.FontBold = False
Me.cmdguardar.FontBold = False
Me.cmdeliminar.FontBold = False
Me.cmdcancelar.FontBold = False
Me.cmdmodificar.FontBold = False
Me.cmdagregar.FontBold = False
Me.cmd_actualiza_precio.FontBold = True

End Sub


Private Sub habilitar_botones(valor As Boolean)
'Me.cmd_materiales.Enabled = valor
'Me.cmd_estado.Enabled = valor
'Me.cmd_liquidacion.Enabled = valor

End Sub

Private Sub txt_cant_inicial_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_cant_actual_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_cant_min_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_cant_max_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_descripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Index <> 6 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 13 Or Index = 6 Or Index = 9 Then
        If ((KeyAscii < 48) Or (KeyAscii > 57)) And (KeyAscii <> 8) Then KeyAscii = 0
    End If
End Sub

Private Sub txt_precio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub


