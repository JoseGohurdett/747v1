Imports System.Data.SqlClient
Imports Entidad
Public Class clsAdicionalDA
    Dim con As New clsConexion 'SE CREA OBJETO DE CLSCONEXION PARA PODER CONECTARNOS A BASE DE DATOS
    Dim query As String 'VARIABLE QUE ALMACENARA EL PROCEDIMIENTO ALMACENDADO A CONSULTAR
    Dim tipoConexion As Integer 'VARIABLE QUE ALMACENARA LA BASE A LA CUAL CONSULTAR
    ' Dim vlSqlParam As New cm 'VARIABLE QUE ALMACENARA LOS PARAMETROS A LA BASE DE DATOS
    Dim Tabla As New DataTable

    Public Sub Insertar(ByVal adic As eAdicional)
        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable
        Dim valido As Boolean

        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Gestion].[pa_IngresaAdicional]"
        cmd.Parameters.AddWithValue("@C_ID", adic.C_ID)
        cmd.Parameters.AddWithValue("@A_NRO", adic.A_NRO)
        cmd.Parameters.AddWithValue("@A_RUT", adic.A_Rut)
        cmd.Parameters.AddWithValue("@A_DV", adic.A_Dv)
        cmd.Parameters.AddWithValue("@A_NOMBRE", adic.A_Nombre)
        cmd.Parameters.AddWithValue("@A_PATERNO", adic.A_Paterno)
        cmd.Parameters.AddWithValue("@A_MATERNO", adic.A_Materno)
        cmd.Parameters.AddWithValue("@A_FECHA_NACIMIENTO", adic.A_FechaNacimiento)
        cmd.Parameters.AddWithValue("@A_SEXO", adic.A_Sexo)
        cmd.Parameters.AddWithValue("@A_TIPO_BENEFICIARIO", adic.A_TipoBeneficiario)
        cmd.CommandText = sql
        valido = con.Ejecutar(cmd, 2)


    End Sub

    Public Function Verificar(ByVal id As String, ByVal rut As String) As Boolean
        'Try

        '    vlSqlParam.Clear()

        '    tipoConexion = 1
        '    query = "[dbo].[Verifica_Adicional]"
        '    vlSqlParam.Add("@cli_id", id, SqlDbType.BigInt)
        '    vlSqlParam.Add("@cli_arut", rut, SqlDbType.Int)
        '    Tabla = con.TraeDatosConP(vlSqlParam, query, tipoConexion)
        '    If Tabla.Rows.Count > 0 Then
        '        Return True
        '    Else
        '        Return False
        '    End If

        '    vlSqlParam.Clear()
        'Catch ex As Exception

        'End Try
    End Function

    Public Sub Eliminar(ByVal _id As String)
        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable
        Dim valido As Boolean

        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Gestion].[pa_EliminarAdicional]"
        cmd.Parameters.AddWithValue("@C_ID", _id)
        cmd.CommandText = sql
        valido = con.Ejecutar(cmd, 2)


    End Sub

    Public Function Carga_Adicional(ByVal id As String) As List(Of eAdicional)
        Dim cmd As New SqlCommand
        Dim sql As String
        Dim dt As New DataTable

        cmd.CommandType = CommandType.StoredProcedure

        sql = "[Gestion].[pa_ListaAdicionales]"
        cmd.CommandText = sql
        cmd.Parameters.AddWithValue("@C_ID", id)
        dt = con.TraeDatos(cmd, 2)

        Dim lstAdicional As New List(Of eAdicional)

        For x As Int16 = 0 To dt.Rows.Count - 1
            Dim row As New eAdicional

            row.A_Rut = IIf(IsDBNull(dt.Rows(x)("A_Rut")), Nothing, dt.Rows(x)("A_Rut"))
            row.A_Dv = IIf(IsDBNull(dt.Rows(x)("A_Dv")), Nothing, dt.Rows(x)("A_Dv"))
            row.A_Nombre = IIf(IsDBNull(dt.Rows(x)("A_Nombre")), Nothing, dt.Rows(x)("A_Nombre"))
            row.A_Paterno = IIf(IsDBNull(dt.Rows(x)("A_Paterno")), Nothing, dt.Rows(x)("A_Paterno"))
            row.A_Materno = IIf(IsDBNull(dt.Rows(x)("A_Materno")), Nothing, dt.Rows(x)("A_Materno"))
            row.A_FechaNacimiento = IIf(IsDBNull(dt.Rows(x)("A_FechaNacimiento")), Nothing, dt.Rows(x)("A_FechaNacimiento"))
            row.A_Sexo = IIf(IsDBNull(dt.Rows(x)("A_Sexo")), Nothing, dt.Rows(x)("A_Sexo"))
            row.A_TipoBeneficiario = IIf(IsDBNull(dt.Rows(x)("A_TipoBeneficiario")), Nothing, dt.Rows(x)("A_TipoBeneficiario"))
            row.A_NombreTipoBeneficiario = IIf(IsDBNull(dt.Rows(x)("NombreTipoBeneficiario")), Nothing, dt.Rows(x)("NombreTipoBeneficiario"))


            lstAdicional.Add(row)
        Next

        Return lstAdicional

    End Function

End Class
