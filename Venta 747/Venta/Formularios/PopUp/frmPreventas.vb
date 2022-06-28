Imports BI
Public Class frmPreventas
    Dim pre As New clsPreVentaBI
    Dim dt As New DataTable

    Private Sub cargaRegistros()
        preVenta = False
        dtgAdicionalesBase.DataSource = Nothing
        dtgAdicionalesBase.Rows.Clear()
        dtgAdicionalesBase.Columns.Clear()
        dt = pre.ListaPreVentas(WS_IDUSUARIO)

        If dt.Rows.Count > 0 Then
            dtgAdicionalesBase.DataSource = dt
            dtgAdicionalesBase.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            For i As Integer = 0 To dtgAdicionalesBase.Rows.Count - 1
                For j As Integer = 0 To dtgAdicionalesBase.Columns.Count - 1
                    If Not dtgAdicionalesBase(j, i).Value Is Nothing Then
                        If dtgAdicionalesBase.Columns(j).Name = "Instalado" Then
                            dtgAdicionalesBase(j, i).ReadOnly = False
                        Else
                            dtgAdicionalesBase(j, i).ReadOnly = True
                        End If
                    End If
                Next
            Next

            Dim cmb As New DataGridViewComboBoxColumn()
            cmb.HeaderText = "Banco"
            cmb.Name = "banco"
            cmb.Items.Add("BANCO DE CHILE")
            cmb.Items.Add("BANCO ESTADO")
            cmb.Items.Add("SCOTIABANK")
            cmb.Items.Add("BCI")
            cmb.Items.Add("CORPBANCA")
            cmb.Items.Add("BANCO SANTANDER")
            cmb.Items.Add("BANCO SECURITY")
            cmb.Items.Add("BANCO FALABELLA")
            cmb.Items.Add("BANCO RIPLEY")
            cmb.Items.Add("BANCO BBVA")
            cmb.Items.Add("BANCO PARIS")
            cmb.Items.Add("BANCO NOVA")
            cmb.Items.Add("BANCO CREDICHILE")
            cmb.Items.Add("ITAÚ")
            dtgAdicionalesBase.Columns.Add(cmb)


            cmb = New DataGridViewComboBoxColumn()
            cmb.HeaderText = "Tarjeta"
            cmb.Name = "tarjeta"
            cmb.Items.Add("VISA")
            cmb.Items.Add("MASTERCARD")
            cmb.Items.Add("PRE PAGO")
            cmb.Items.Add("AMERICAN EXPRESS")
            dtgAdicionalesBase.Columns.Add(cmb)
        End If
    End Sub

    Private Sub frmPreventas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cargaRegistros()
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        For i As Integer = 0 To dtgAdicionalesBase.Rows.Count - 1
            For j As Integer = 0 To dtgAdicionalesBase.Columns.Count - 1
                If dtgAdicionalesBase.Item("Instalado", i).Value Then
                    If Not dtgAdicionalesBase(j, i).Value Is Nothing Then
                        Dim idCliente As Int64 = dtgAdicionalesBase.Item("idCliente", i).Value.ToString()
                        Dim banco As String = dtgAdicionalesBase.Item("banco", i).Value.ToString()
                        Dim tarjeta As String = dtgAdicionalesBase.Item("tarjeta", i).Value.ToString()
                        If banco = "" Then
                            MsgBox("Debe seleccionar el banco")
                            Exit Sub
                        End If
                        If tarjeta = "" Then
                            MsgBox("Debe seleccionar la tarjeta")
                            Exit Sub
                        End If
                        pre.ActualizaPreVenta(idCliente, banco, tarjeta)
                    End If
                End If
            Next
        Next

        cargaRegistros()
    End Sub


    Private Sub dtgAdicionalesBase_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dtgAdicionalesBase.CellContentDoubleClick
        If e.RowIndex >= 0 Then
            perfil = "Regrabador"
            preVenta = True
            GesId = dtgAdicionalesBase.Rows(e.RowIndex).Cells("idCliente").Value
            estado_perfil = True
            Me.Hide()
            frmVenta.Show()
        End If
    End Sub

    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        preVenta = False
        Me.Hide()
        frmLogin.Show()
    End Sub
End Class