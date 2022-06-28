Imports System.Data.SqlClient
Imports Entidad

Public Class clsPreVentaDA
    Dim con As New clsConexion

    Public Function ListaPreVentas(ByVal _ejecutivo As String) As DataTable
        Dim cmd As New SqlCommand
        Dim dt As New DataTable

        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = "[Gestion].[pa_ListaPreVenta]"
        cmd.Parameters.AddWithValue("@ejecutivo", _ejecutivo)
        dt = con.TraeDatos(cmd, 2)

        Return dt
    End Function

    Public Sub ActualizaPreVenta(ByVal _idCliente As Int64, ByVal _banco As String, ByVal _tarjeta As String)
        Dim cmd As New SqlCommand
        Dim dt As New DataTable

        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = "[Gestion].[pa_ActualizaPreVenta]"
        cmd.Parameters.AddWithValue("@idCliente", _idCliente)
        cmd.Parameters.AddWithValue("@banco", _banco)
        cmd.Parameters.AddWithValue("@tarjeta", _tarjeta)
        con.Ejecutar(cmd, 2)

    End Sub

End Class
