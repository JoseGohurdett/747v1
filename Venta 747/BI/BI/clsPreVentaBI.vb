Imports Dato
Public Class clsPreVentaBI
    Dim pre As New clsPreVentaDA
    Function ListaPreVentas(ByVal _ejecutivo As String) As DataTable
        Return pre.ListaPreVentas(_ejecutivo)
    End Function

    Public Sub ActualizaPreVenta(ByVal _idCliente As Int64, ByVal _banco As String, ByVal _tarjeta As String)
        pre.ActualizaPreVenta(_idCliente, _banco, _tarjeta)
    End Sub
End Class
