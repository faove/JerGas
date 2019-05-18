VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form rpt_resumen_mensual_ventas 
   Caption         =   "RESUMEN MENSUAL DE VENTAS"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "rpt_resumen_mensual_ventas.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      lastProp        =   500
      _cx             =   10231
      _cy             =   12347
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "rpt_resumen_mensual_ventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crt_resumen_mensual_ventas

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
CRViewer91.ReportSource = Report
Report.DiscardSavedData

If IsNull(frm_resumen_mensual_ventas.DTPicker_desde.Value) Or IsNull(frm_resumen_mensual_ventas.DTPicker_hasta.Value) Then
        MsgBox "Por favor, los valores de fecha desde y fecha hasta no pueden ser nulos. Gracias", vbCritical, "JerGas, C.A."
        Exit Sub
    End If
    If frm_resumen_mensual_ventas.DTPicker_desde.Value > frm_resumen_mensual_ventas.DTPicker_hasta.Value Then
        MsgBox "Por favor, verifique que la fecha incial no sea mayor que la fecha final. Gracias", vbCritical, "JerGas, C.A."
        Exit Sub
    End If
    SELECCION = "({tbl_resumen_mensual_ventas.Fecha_venta} >= #" & Format(frm_resumen_mensual_ventas.DTPicker_desde.Value, "mm/dd/yyyy") & "#  and {tbl_resumen_mensual_ventas.Fecha_venta} <= #" & Format(frm_resumen_mensual_ventas.DTPicker_hasta.Value, "mm/dd/yyyy") & "#)"
 '   SELECCION = "({resumen_pedidos.Fecha_pedido} >= #" & Format(frm_resumen_mensual_ventas.DTPicker_desde.Value, "mm/dd/yyyy") & "#  and {resumen_pedidos.Fecha_pedido} <= #" & Format(frm_resumen_mensual_ventas.DTPicker_hasta.Value, "mm/dd/yyyy") & "#)"

    Report.RecordSelectionFormula = SELECCION






CRViewer91.ViewReport
Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
CRViewer91.Top = 0
CRViewer91.Left = 0
CRViewer91.Height = ScaleHeight
CRViewer91.Width = ScaleWidth

End Sub
