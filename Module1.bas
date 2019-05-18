Attribute VB_Name = "Module1"
Public cn As ADODB.Connection
Public Rdsliq As ADODB.Recordset
Public Rdscontrol As ADODB.Recordset
Public Rdsfactura As ADODB.Recordset
Public Rdsped As ADODB.Recordset
Public I As Byte
Public Usuario, user_name, user_grupo
Public enlace1, enlace2


Public Sub actualizar_cn(PRODRIVER As String)
    
    Set cn = New ADODB.Connection
    
    cn.CommandTimeout = 180
    
'    cn.Open "Driver={" & PRODRIVER & "};Server=SOCASV;Uid=sa;Pwd=;Database=ALCALSIS"
    cn.Open "DSN=gergas"
'    cn.Open "Driver={SQL Server};Server=G6T6I0;Uid=nelson;Pwd=nelson;Database=ALCALSIS"

'    cn.Open PRODRIVER
    
    'MsgBox "Conexion a SqlServer Exitosa."
    

End Sub

Public Sub Mover_der(Objeto As Object, Obj_mover As Object, Separar As Single)

Dim Izq, Ancho_obj, Ancho As Single
    
    Ancho_obj = Obj_mover.Width
    Ancho = Objeto.ScaleWidth
    Izq = Ancho - Ancho_obj
    Obj_mover.Move Izq - Separar

End Sub
Public Sub Mover_centrado(Objeto As Object, Obj_mover As Object)

Dim Izq, Ancho_obj, Ancho As Single
    
    Ancho_obj = Obj_mover.Width
    Ancho = Objeto.ScaleWidth
    Izq = (Ancho - Ancho_obj) / 2
    Obj_mover.Move Izq

End Sub

Public Function FGNRO_LIQ()
    
    ABRIR_RdsLiq
    Gnro_Liquida = Val(Rdsliq!Nro_liquida_ult)
    Gnro_Liquida = Gnro_Liquida + 1
      FGNRO_LIQ = Gnro_Liquida
    Rdsliq!Nro_liquida_ult = Gnro_Liquida
    Rdsliq.Update
    Rdsliq.Close
End Function

Public Function FGNRO_CONTROL()
    
    ABRIR_Rdscontrol
    Gnro_control = Val(Rdscontrol!id_control)
    Gnro_control = Gnro_control + 1
      FGNRO_CONTROL = Gnro_control
    Rdscontrol!id_control = Gnro_control
    Rdscontrol.Update
    Rdscontrol.Close
End Function

Public Function FGNRO_FACTURA()
    
    ABRIR_Rdsfactura
    Gnro_factura = Val(Rdsfactura!id_factura)
    Gnro_factura = Gnro_factura + 1
      FGNRO_FACTURA = Gnro_factura
    Rdsfactura!id_factura = Gnro_factura
    Rdsfactura.Update
    Rdsfactura.Close
End Function

Private Function ABRIR_RdsLiq()

Set Rdsliq = New ADODB.Recordset

    Rdsliq.CursorType = adOpenKeyset
    Rdsliq.LockType = adLockPessimistic 'desbloquea el objeto recordset
    Rdsliq.Open "select * from tbl_control_procesos", cn

End Function

Private Function ABRIR_Rdsfactura()

Set Rdsfactura = New ADODB.Recordset

    Rdsfactura.CursorType = adOpenKeyset
    Rdsfactura.LockType = adLockPessimistic 'desbloquea el objeto recordset
    Rdsfactura.Open "select * from tbl_control_procesos", cn

End Function

Private Function ABRIR_Rdscontrol()

Set Rdscontrol = New ADODB.Recordset

    Rdscontrol.CursorType = adOpenKeyset
    Rdscontrol.LockType = adLockPessimistic 'desbloquea el objeto recordset
    Rdscontrol.Open "select * from tbl_control_procesos", cn

End Function

Private Function ABRIR_Rdsped()

Set Rdsped = New ADODB.Recordset

    Rdsped.CursorType = adOpenKeyset
    Rdsped.LockType = adLockPessimistic 'desbloquea el objeto recordset
    Rdsped.Open "select * from tbl_pedidos", cn

End Function

Sub limpia(forma As Form)
 For I = 0 To forma.Controls.Count - 1
    If TypeOf forma.Controls(I) Is TextBox Then
           forma.Controls(I) = ""
    End If
    If TypeOf forma.Controls(I) Is ComboBox Then
           forma.Controls(I) = "Seleccione"
    End If
 Next I
End Sub

Public Function FGNRO_LIQ_RESTA_PEDIDO()
    
    ABRIR_RdsLiq
    Gnro_Liquida = Val(Rdsliq!Nro_liquida_ult)
    Gnro_Liquida = Gnro_Liquida - 1
       FGNRO_LIQ_RESTA_PEDIDO = Gnro_Liquida
    Rdsliq!Nro_liquida_ult = Gnro_Liquida
    Rdsliq.Update
    Rdsliq.Close
End Function

Public Function FGNRO_LIQ_RESTA_FACTURA()
    
    ABRIR_Rdsfactura
    Gnro_factura = Val(Rdsfactura!id_factura)
    Gnro_factura = Gnro_factura - 1
       FGNRO_LIQ_RESTA_FACTURA = Gnro_factura
    Rdsfactura!id_factura = Gnro_factura
    Rdsfactura.Update
    Rdsfactura.Close
End Function

Public Function FGNRO_LIQ_RESTA_CONTROL()
    
    ABRIR_Rdscontrol
    Gnro_control = Val(Rdscontrol!id_control)
    Gnro_control = Gnro_control - 1
      FGNRO_LIQ_RESTA_CONTROL = Gnro_control
    Rdscontrol!id_control = Gnro_control
    Rdscontrol.Update
    Rdscontrol.Close
End Function

