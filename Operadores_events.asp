<%
'BindEvents Method @1-52861504
Sub BindEvents(Level)
    If Level="Page" Then
    Else
        Set Operadores1.Fecha_de_Alta.CCSEvents("BeforeShow") = GetRef("Operadores1_Fecha_de_Alta_BeforeShow")
    End If
End Sub
'End BindEvents Method

'Operadores1_Fecha_de_Alta_BeforeShow @24-CB24AEAE
Function Operadores1_Fecha_de_Alta_BeforeShow(Sender)
'End Operadores1_Fecha_de_Alta_BeforeShow

'Close Operadores1_Fecha_de_Alta_BeforeShow @24-54C34B28
End Function
'End Close Operadores1_Fecha_de_Alta_BeforeShow


%>
