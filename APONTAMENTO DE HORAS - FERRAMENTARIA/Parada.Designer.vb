<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Parada
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Parada))
        Me.txt_motivo = New System.Windows.Forms.Label()
        Me.cbo_motivo = New System.Windows.Forms.ComboBox()
        Me.btn_parar = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txt_motivo
        '
        Me.txt_motivo.AutoSize = True
        Me.txt_motivo.Font = New System.Drawing.Font("Arial Narrow", 15.0!, System.Drawing.FontStyle.Bold)
        Me.txt_motivo.Location = New System.Drawing.Point(120, 102)
        Me.txt_motivo.Name = "txt_motivo"
        Me.txt_motivo.Size = New System.Drawing.Size(0, 24)
        Me.txt_motivo.TabIndex = 1
        '
        'cbo_motivo
        '
        Me.cbo_motivo.Font = New System.Drawing.Font("Arial Narrow", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbo_motivo.FormattingEnabled = True
        Me.cbo_motivo.Location = New System.Drawing.Point(57, 131)
        Me.cbo_motivo.Name = "cbo_motivo"
        Me.cbo_motivo.Size = New System.Drawing.Size(269, 31)
        Me.cbo_motivo.TabIndex = 1
        '
        'btn_parar
        '
        Me.btn_parar.BackColor = System.Drawing.Color.Transparent
        Me.btn_parar.FlatAppearance.BorderColor = System.Drawing.Color.OrangeRed
        Me.btn_parar.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.btn_parar.FlatAppearance.MouseOverBackColor = System.Drawing.Color.OrangeRed
        Me.btn_parar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_parar.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_parar.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn_parar.Location = New System.Drawing.Point(57, 171)
        Me.btn_parar.Name = "btn_parar"
        Me.btn_parar.Size = New System.Drawing.Size(269, 31)
        Me.btn_parar.TabIndex = 83
        Me.btn_parar.Text = "PARAR"
        Me.btn_parar.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(126, 106)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(130, 18)
        Me.Label1.TabIndex = 85
        Me.Label1.Text = "Motivo da parada"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.Ordem_de_Servico.My.Resources.Resources.logo
        Me.PictureBox1.Location = New System.Drawing.Point(84, 51)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(211, 26)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 84
        Me.PictureBox1.TabStop = False
        '
        'Parada
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.ClientSize = New System.Drawing.Size(382, 247)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.btn_parar)
        Me.Controls.Add(Me.cbo_motivo)
        Me.Controls.Add(Me.txt_motivo)
        Me.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "Parada"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Parada"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txt_motivo As Label
    Friend WithEvents cbo_motivo As ComboBox
    Friend WithEvents btn_parar As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents Label1 As Label
End Class
