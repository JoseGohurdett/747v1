Imports System.Collections.Generic
Imports Dato
Imports Entidad
Public Class clsParentescoCampaniaBI

    Public Function BuscarParentescoPorId(ByVal _idCRM As Integer, ByVal _tipoPersona As Integer) As List(Of eTipoParentesco)
        Dim daParentescoCampania As New clsParentescoCampaniaDA
        Return daParentescoCampania.BuscaParentescoPorId(_idCRM, _tipoPersona)
    End Function
End Class
