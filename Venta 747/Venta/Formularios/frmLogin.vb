Option Strict Off
Option Explicit On

Imports Entidad
Imports BI

Friend Class frmLogin
    Inherits System.Windows.Forms.Form
    Dim opc As Short
    '*******metodo al presionar boton aceptar de boton login****************
    Private Sub btnAceptar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnAceptar.Click
        opc = 1
        Dim valido As Boolean
        Dim bigene As New clsGeneralBI
        Dim biUsuario As New clsUsuarioBI
        WS_IDUSUARIO = TxtUsuarioLog.Text
        If CheckBoxLogin.Checked = True Then ' si checkbox esta marcado es para entrar como regrabador
            valido = biUsuario.Validar_Reg(WS_IDUSUARIO) ' validamos si el usuario tiene  el perfil de ejecutivo regrabador
            If valido = True Then
                perfil = "Regrabador"
            Else
                valido = biUsuario.Validar_Reg(WS_IDUSUARIO) ' validamos si el usuario tiene  el perfil de sistemas
                If valido = True Then
                    perfil = "Regrabador"
                End If
            End If
            If valido = False Then
                MsgBox("El usuario no tiene permisos como Ejecutivo de Regrabación", MsgBoxStyle.Exclamation, "Atención")
                End
            End If

        ElseIf CheckBoxLogin.Checked = False Then ' si checkbox no esta marcado entra como ventas
            valido = biUsuario.Validar_user(WS_IDUSUARIO) ' validamos el usuario para ventas
            If valido = True Then
                perfil = "Ejecutivo"
            Else
                MsgBox("El usuario no tiene permisos como Ejecutivo de Ventas", MsgBoxStyle.Exclamation, "Atención")
                Exit Sub
                'End
            End If
        End If

        If chkPreVenta.Checked Then
            frmPreventas.ShowDialog()
            'Exit Sub   
        End If


        If ConectaNeotel((txtusuxlite.Text)) Then
            ModGeneral.conectarTelefonia()
            Logear_Usuario(WS_IDUSUARIO, 1)
            frmVenta.Show()
            Me.Close()
        End If

    End Sub


    Private Sub frmLogin_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Me.Text = "NEW Versión: " & My.Application.Info.Version.Major.ToString & "." & My.Application.Info.Version.Minor.ToString _
      & "." & My.Application.Info.Version.Revision.ToString & "-" & csNombreAplicacion
        'TxtUsuarioLog.Text = Trim(WS_IDUSUARIO)
        'TXT_USUARIO.Text = WS_IDUSUARIO
        If db_central = 4 Then
            fraxlite.Visible = True
        End If
        opc = 0
    End Sub

    Public Function ConectaNeotel(ByRef Usuario As String) As Boolean
        If Usuario = "" Then
            MsgBox("Debe ingresar usuario XLite", MsgBoxStyle.Exclamation, "Atención")
            txtusuxlite.Focus()
        Else
            'vpPosicion = New clsSoapNeo
            If vpPosicion.CargarPosicion(Usuario) <> "" Then
                If vpPosicion.Usuario = Usuario Then
                    ConectaNeotel = True
                Else
                    MsgBox("El usuario XLite no corresponde", MsgBoxStyle.Exclamation, "Atención")
                End If
            End If
        End If
    End Function

    Private Sub frmLogin_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If opc = 0 Then
            End
        End If
    End Sub

    Private Sub TxtUsuarioLog_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtUsuarioLog.KeyPress
        If IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtusuxlite_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtusuxlite.KeyPress
        If (Not IsNumeric(e.KeyChar)) Then
            If (e.KeyChar <> vbBack) Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Me.Close()
        End
    End Sub

    Private Sub chkPreVenta_CheckedChanged(sender As Object, e As EventArgs) Handles chkPreVenta.CheckedChanged

    End Sub
End Class