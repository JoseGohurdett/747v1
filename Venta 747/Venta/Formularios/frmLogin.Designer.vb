<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmLogin
#Region "Código generado por el Diseñador de Windows Forms "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'Llamada necesaria para el Diseñador de Windows Forms.
		InitializeComponent()
	End Sub
	'Form invalida a Dispose para limpiar la lista de componentes.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Requerido por el Diseñador de Windows Forms
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents txtusuxlite As System.Windows.Forms.TextBox
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents fraxlite As System.Windows.Forms.Panel
	Public WithEvents TxtUsuarioLog As System.Windows.Forms.TextBox
	Public WithEvents btnAceptar As System.Windows.Forms.Button
	Public WithEvents LblBienvenida As System.Windows.Forms.Label
    Public WithEvents TXT_USUARIO As System.Windows.Forms.Label
	'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
	'Se puede modificar mediante el Diseñador de Windows Forms.
	'No lo modifique con el editor de código.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLogin))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.fraxlite = New System.Windows.Forms.Panel()
        Me.txtusuxlite = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TxtUsuarioLog = New System.Windows.Forms.TextBox()
        Me.btnAceptar = New System.Windows.Forms.Button()
        Me.LblBienvenida = New System.Windows.Forms.Label()
        Me.TXT_USUARIO = New System.Windows.Forms.Label()
        Me.CheckBoxLogin = New System.Windows.Forms.CheckBox()
        Me.Image1 = New System.Windows.Forms.PictureBox()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.chkPreVenta = New System.Windows.Forms.CheckBox()
        Me.fraxlite.SuspendLayout()
        CType(Me.Image1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'fraxlite
        '
        Me.fraxlite.BackColor = System.Drawing.Color.White
        Me.fraxlite.Controls.Add(Me.txtusuxlite)
        Me.fraxlite.Controls.Add(Me.Label1)
        Me.fraxlite.Cursor = System.Windows.Forms.Cursors.Default
        Me.fraxlite.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraxlite.Location = New System.Drawing.Point(25, 315)
        Me.fraxlite.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.fraxlite.Name = "fraxlite"
        Me.fraxlite.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraxlite.Size = New System.Drawing.Size(396, 62)
        Me.fraxlite.TabIndex = 4
        Me.fraxlite.Text = "Frame1"
        Me.fraxlite.Visible = False
        '
        'txtusuxlite
        '
        Me.txtusuxlite.AcceptsReturn = True
        Me.txtusuxlite.BackColor = System.Drawing.SystemColors.Window
        Me.txtusuxlite.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtusuxlite.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtusuxlite.Location = New System.Drawing.Point(171, 20)
        Me.txtusuxlite.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtusuxlite.MaxLength = 0
        Me.txtusuxlite.Name = "txtusuxlite"
        Me.txtusuxlite.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtusuxlite.Size = New System.Drawing.Size(107, 22)
        Me.txtusuxlite.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.White
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(4, 23)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(159, 21)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Ingrese usuario XLite"
        '
        'TxtUsuarioLog
        '
        Me.TxtUsuarioLog.AcceptsReturn = True
        Me.TxtUsuarioLog.BackColor = System.Drawing.SystemColors.Window
        Me.TxtUsuarioLog.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtUsuarioLog.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxtUsuarioLog.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TxtUsuarioLog.Location = New System.Drawing.Point(164, 284)
        Me.TxtUsuarioLog.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TxtUsuarioLog.MaxLength = 0
        Me.TxtUsuarioLog.Name = "TxtUsuarioLog"
        Me.TxtUsuarioLog.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxtUsuarioLog.Size = New System.Drawing.Size(256, 22)
        Me.TxtUsuarioLog.TabIndex = 2
        '
        'btnAceptar
        '
        Me.btnAceptar.BackColor = System.Drawing.SystemColors.Control
        Me.btnAceptar.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnAceptar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnAceptar.Location = New System.Drawing.Point(144, 390)
        Me.btnAceptar.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnAceptar.Name = "btnAceptar"
        Me.btnAceptar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnAceptar.Size = New System.Drawing.Size(133, 41)
        Me.btnAceptar.TabIndex = 0
        Me.btnAceptar.Text = "Ingresar"
        Me.btnAceptar.UseVisualStyleBackColor = False
        '
        'LblBienvenida
        '
        Me.LblBienvenida.BackColor = System.Drawing.Color.White
        Me.LblBienvenida.Cursor = System.Windows.Forms.Cursors.Default
        Me.LblBienvenida.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblBienvenida.Location = New System.Drawing.Point(25, 288)
        Me.LblBienvenida.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LblBienvenida.Name = "LblBienvenida"
        Me.LblBienvenida.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LblBienvenida.Size = New System.Drawing.Size(140, 31)
        Me.LblBienvenida.TabIndex = 3
        Me.LblBienvenida.Text = "Usuario Logeado"
        '
        'TXT_USUARIO
        '
        Me.TXT_USUARIO.BackColor = System.Drawing.Color.White
        Me.TXT_USUARIO.Cursor = System.Windows.Forms.Cursors.Default
        Me.TXT_USUARIO.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.TXT_USUARIO.Location = New System.Drawing.Point(85, 315)
        Me.TXT_USUARIO.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.TXT_USUARIO.Name = "TXT_USUARIO"
        Me.TXT_USUARIO.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TXT_USUARIO.Size = New System.Drawing.Size(271, 16)
        Me.TXT_USUARIO.TabIndex = 1
        Me.TXT_USUARIO.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'CheckBoxLogin
        '
        Me.CheckBoxLogin.AutoSize = True
        Me.CheckBoxLogin.Location = New System.Drawing.Point(33, 414)
        Me.CheckBoxLogin.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.CheckBoxLogin.Name = "CheckBoxLogin"
        Me.CheckBoxLogin.Size = New System.Drawing.Size(18, 17)
        Me.CheckBoxLogin.TabIndex = 6
        Me.CheckBoxLogin.UseVisualStyleBackColor = True
        '
        'Image1
        '
        Me.Image1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Image1.Image = Global.APP_747.My.Resources.Resources.Imagen_Cordial_Phone1
        Me.Image1.Location = New System.Drawing.Point(73, 2)
        Me.Image1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Image1.Name = "Image1"
        Me.Image1.Size = New System.Drawing.Size(283, 245)
        Me.Image1.TabIndex = 5
        Me.Image1.TabStop = False
        '
        'btnSalir
        '
        Me.btnSalir.BackColor = System.Drawing.SystemColors.Control
        Me.btnSalir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnSalir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnSalir.Location = New System.Drawing.Point(177, 457)
        Me.btnSalir.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnSalir.Name = "btnSalir"
        Me.btnSalir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnSalir.Size = New System.Drawing.Size(83, 26)
        Me.btnSalir.TabIndex = 7
        Me.btnSalir.Text = "Salir"
        Me.btnSalir.UseVisualStyleBackColor = False
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.White
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(4, 434)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(104, 28)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Recuperacion"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.White
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(316, 414)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(104, 28)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "PreVenta"
        '
        'chkPreVenta
        '
        Me.chkPreVenta.AutoSize = True
        Me.chkPreVenta.Location = New System.Drawing.Point(345, 394)
        Me.chkPreVenta.Margin = New System.Windows.Forms.Padding(4)
        Me.chkPreVenta.Name = "chkPreVenta"
        Me.chkPreVenta.Size = New System.Drawing.Size(18, 17)
        Me.chkPreVenta.TabIndex = 9
        Me.chkPreVenta.UseVisualStyleBackColor = True
        '
        'frmLogin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.HighlightText
        Me.ClientSize = New System.Drawing.Size(433, 497)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.chkPreVenta)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnSalir)
        Me.Controls.Add(Me.CheckBoxLogin)
        Me.Controls.Add(Me.fraxlite)
        Me.Controls.Add(Me.TxtUsuarioLog)
        Me.Controls.Add(Me.btnAceptar)
        Me.Controls.Add(Me.LblBienvenida)
        Me.Controls.Add(Me.Image1)
        Me.Controls.Add(Me.TXT_USUARIO)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(3, 27)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmLogin"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " "
        Me.fraxlite.ResumeLayout(False)
        Me.fraxlite.PerformLayout()
        CType(Me.Image1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents Image1 As System.Windows.Forms.PictureBox
    Friend WithEvents CheckBoxLogin As System.Windows.Forms.CheckBox
    Public WithEvents btnSalir As Button
    Public WithEvents Label2 As Label
    Public WithEvents Label3 As Label
    Friend WithEvents chkPreVenta As CheckBox
#End Region
End Class