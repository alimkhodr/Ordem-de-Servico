<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Desenho
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Desenho))
        Me.txt_motivo = New System.Windows.Forms.Label()
        Me.btn_imprimir = New System.Windows.Forms.Button()
        Me.btn_procurar = New System.Windows.Forms.Button()
        Me.txt_caminhadaimagem = New System.Windows.Forms.Label()
        Me.PictureBox = New System.Windows.Forms.PictureBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        CType(Me.PictureBox, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txt_motivo
        '
        Me.txt_motivo.AutoSize = True
        Me.txt_motivo.Font = New System.Drawing.Font("Arial Narrow", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_motivo.Location = New System.Drawing.Point(56, 97)
        Me.txt_motivo.Name = "txt_motivo"
        Me.txt_motivo.Size = New System.Drawing.Size(169, 23)
        Me.txt_motivo.TabIndex = 1
        Me.txt_motivo.Text = "Selecione o desenho"
        '
        'btn_imprimir
        '
        Me.btn_imprimir.BackColor = System.Drawing.Color.Transparent
        Me.btn_imprimir.FlatAppearance.BorderColor = System.Drawing.Color.OrangeRed
        Me.btn_imprimir.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.btn_imprimir.FlatAppearance.MouseOverBackColor = System.Drawing.Color.OrangeRed
        Me.btn_imprimir.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_imprimir.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_imprimir.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn_imprimir.Location = New System.Drawing.Point(192, 302)
        Me.btn_imprimir.Name = "btn_imprimir"
        Me.btn_imprimir.Size = New System.Drawing.Size(133, 33)
        Me.btn_imprimir.TabIndex = 83
        Me.btn_imprimir.Text = "IMPRIMIR"
        Me.btn_imprimir.UseVisualStyleBackColor = False
        '
        'btn_procurar
        '
        Me.btn_procurar.BackColor = System.Drawing.Color.Transparent
        Me.btn_procurar.FlatAppearance.BorderColor = System.Drawing.Color.OrangeRed
        Me.btn_procurar.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.btn_procurar.FlatAppearance.MouseOverBackColor = System.Drawing.Color.OrangeRed
        Me.btn_procurar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_procurar.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_procurar.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn_procurar.Location = New System.Drawing.Point(60, 302)
        Me.btn_procurar.Name = "btn_procurar"
        Me.btn_procurar.Size = New System.Drawing.Size(126, 33)
        Me.btn_procurar.TabIndex = 85
        Me.btn_procurar.Text = "PROCURAR"
        Me.btn_procurar.UseVisualStyleBackColor = False
        '
        'txt_caminhadaimagem
        '
        Me.txt_caminhadaimagem.AutoSize = True
        Me.txt_caminhadaimagem.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_caminhadaimagem.Location = New System.Drawing.Point(57, 270)
        Me.txt_caminhadaimagem.Name = "txt_caminhadaimagem"
        Me.txt_caminhadaimagem.Size = New System.Drawing.Size(75, 16)
        Me.txt_caminhadaimagem.TabIndex = 87
        Me.txt_caminhadaimagem.Text = "imagem.jpg"
        Me.txt_caminhadaimagem.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'PictureBox
        '
        Me.PictureBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox.Location = New System.Drawing.Point(60, 123)
        Me.PictureBox.Name = "PictureBox"
        Me.PictureBox.Size = New System.Drawing.Size(264, 144)
        Me.PictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox.TabIndex = 86
        Me.PictureBox.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.Ordem_de_Servico.My.Resources.Resources.logo
        Me.PictureBox1.Location = New System.Drawing.Point(85, 49)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(211, 26)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 84
        Me.PictureBox1.TabStop = False
        '
        'Desenho
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.ClientSize = New System.Drawing.Size(382, 376)
        Me.Controls.Add(Me.txt_caminhadaimagem)
        Me.Controls.Add(Me.PictureBox)
        Me.Controls.Add(Me.btn_procurar)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.btn_imprimir)
        Me.Controls.Add(Me.txt_motivo)
        Me.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "Desenho"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Desenho"
        CType(Me.PictureBox, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txt_motivo As Label
    Friend WithEvents btn_imprimir As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents btn_procurar As Button
    Friend WithEvents PictureBox As PictureBox
    Friend WithEvents txt_caminhadaimagem As Label
End Class
