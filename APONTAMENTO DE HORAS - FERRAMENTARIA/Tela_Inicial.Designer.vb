<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Tela_Inicial
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Tela_Inicial))
        Me.txt_fer = New System.Windows.Forms.TextBox()
        Me.btn_relatorio = New System.Windows.Forms.Button()
        Me.btn_servico = New System.Windows.Forms.Button()
        Me.btn_maq = New System.Windows.Forms.Button()
        Me.btn_user = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btn_minimizar = New System.Windows.Forms.Button()
        Me.btn_expandir = New System.Windows.Forms.Button()
        Me.btn_fechar = New System.Windows.Forms.Button()
        Me.btn_home = New System.Windows.Forms.Button()
        Me.btn_os = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.txt_situacao = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Panel1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txt_fer
        '
        Me.txt_fer.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt_fer.Font = New System.Drawing.Font("Arial Narrow", 18.0!)
        Me.txt_fer.Location = New System.Drawing.Point(50, 117)
        Me.txt_fer.Margin = New System.Windows.Forms.Padding(6)
        Me.txt_fer.Name = "txt_fer"
        Me.txt_fer.Size = New System.Drawing.Size(237, 28)
        Me.txt_fer.TabIndex = 6
        '
        'btn_relatorio
        '
        Me.btn_relatorio.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.btn_relatorio.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.btn_relatorio.FlatAppearance.BorderSize = 0
        Me.btn_relatorio.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.btn_relatorio.FlatAppearance.MouseOverBackColor = System.Drawing.Color.OrangeRed
        Me.btn_relatorio.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_relatorio.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Bold)
        Me.btn_relatorio.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn_relatorio.Location = New System.Drawing.Point(613, 1)
        Me.btn_relatorio.Margin = New System.Windows.Forms.Padding(6)
        Me.btn_relatorio.Name = "btn_relatorio"
        Me.btn_relatorio.Size = New System.Drawing.Size(114, 51)
        Me.btn_relatorio.TabIndex = 3
        Me.btn_relatorio.Text = "Relatório"
        Me.btn_relatorio.UseVisualStyleBackColor = False
        '
        'btn_servico
        '
        Me.btn_servico.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.btn_servico.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.btn_servico.FlatAppearance.BorderSize = 0
        Me.btn_servico.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.btn_servico.FlatAppearance.MouseOverBackColor = System.Drawing.Color.OrangeRed
        Me.btn_servico.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_servico.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Bold)
        Me.btn_servico.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn_servico.Location = New System.Drawing.Point(410, 1)
        Me.btn_servico.Margin = New System.Windows.Forms.Padding(6)
        Me.btn_servico.Name = "btn_servico"
        Me.btn_servico.Size = New System.Drawing.Size(191, 51)
        Me.btn_servico.TabIndex = 2
        Me.btn_servico.Text = "Apontar Serviço"
        Me.btn_servico.UseVisualStyleBackColor = False
        '
        'btn_maq
        '
        Me.btn_maq.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.btn_maq.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.btn_maq.FlatAppearance.BorderSize = 0
        Me.btn_maq.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.btn_maq.FlatAppearance.MouseOverBackColor = System.Drawing.Color.OrangeRed
        Me.btn_maq.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_maq.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Bold)
        Me.btn_maq.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn_maq.Location = New System.Drawing.Point(739, 1)
        Me.btn_maq.Margin = New System.Windows.Forms.Padding(6)
        Me.btn_maq.Name = "btn_maq"
        Me.btn_maq.Size = New System.Drawing.Size(115, 51)
        Me.btn_maq.TabIndex = 4
        Me.btn_maq.Text = " Máquina"
        Me.btn_maq.UseVisualStyleBackColor = False
        '
        'btn_user
        '
        Me.btn_user.BackColor = System.Drawing.Color.OrangeRed
        Me.btn_user.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btn_user.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(49, Byte), Integer), CType(CType(49, Byte), Integer))
        Me.btn_user.FlatAppearance.BorderSize = 0
        Me.btn_user.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(49, Byte), Integer), CType(CType(49, Byte), Integer))
        Me.btn_user.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(49, Byte), Integer), CType(CType(49, Byte), Integer))
        Me.btn_user.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_user.Font = New System.Drawing.Font("Arial Narrow", 16.0!, System.Drawing.FontStyle.Bold)
        Me.btn_user.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn_user.Location = New System.Drawing.Point(51, 1)
        Me.btn_user.Margin = New System.Windows.Forms.Padding(6)
        Me.btn_user.Name = "btn_user"
        Me.btn_user.Size = New System.Drawing.Size(76, 51)
        Me.btn_user.TabIndex = 84
        Me.btn_user.UseVisualStyleBackColor = False
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Panel1.Controls.Add(Me.btn_minimizar)
        Me.Panel1.Controls.Add(Me.btn_expandir)
        Me.Panel1.Controls.Add(Me.btn_fechar)
        Me.Panel1.Controls.Add(Me.btn_user)
        Me.Panel1.Controls.Add(Me.btn_home)
        Me.Panel1.Controls.Add(Me.btn_os)
        Me.Panel1.Controls.Add(Me.btn_maq)
        Me.Panel1.Controls.Add(Me.btn_servico)
        Me.Panel1.Controls.Add(Me.btn_relatorio)
        Me.Panel1.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Panel1.Location = New System.Drawing.Point(-1, -1)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1023, 76)
        Me.Panel1.TabIndex = 82
        '
        'btn_minimizar
        '
        Me.btn_minimizar.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_minimizar.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn_minimizar.BackgroundImage = Global.Ordem_de_Servico.My.Resources.Resources._6
        Me.btn_minimizar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btn_minimizar.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(49, Byte), Integer), CType(CType(49, Byte), Integer))
        Me.btn_minimizar.FlatAppearance.BorderSize = 0
        Me.btn_minimizar.FlatAppearance.MouseDownBackColor = System.Drawing.SystemColors.AppWorkspace
        Me.btn_minimizar.FlatAppearance.MouseOverBackColor = System.Drawing.SystemColors.AppWorkspace
        Me.btn_minimizar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_minimizar.Font = New System.Drawing.Font("Arial Narrow", 16.0!, System.Drawing.FontStyle.Bold)
        Me.btn_minimizar.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn_minimizar.Location = New System.Drawing.Point(879, 1)
        Me.btn_minimizar.Margin = New System.Windows.Forms.Padding(6)
        Me.btn_minimizar.Name = "btn_minimizar"
        Me.btn_minimizar.Size = New System.Drawing.Size(42, 39)
        Me.btn_minimizar.TabIndex = 91
        Me.btn_minimizar.UseVisualStyleBackColor = False
        '
        'btn_expandir
        '
        Me.btn_expandir.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_expandir.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn_expandir.BackgroundImage = Global.Ordem_de_Servico.My.Resources.Resources._7
        Me.btn_expandir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btn_expandir.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(49, Byte), Integer), CType(CType(49, Byte), Integer))
        Me.btn_expandir.FlatAppearance.BorderSize = 0
        Me.btn_expandir.FlatAppearance.MouseDownBackColor = System.Drawing.SystemColors.AppWorkspace
        Me.btn_expandir.FlatAppearance.MouseOverBackColor = System.Drawing.SystemColors.AppWorkspace
        Me.btn_expandir.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_expandir.Font = New System.Drawing.Font("Arial Narrow", 16.0!, System.Drawing.FontStyle.Bold)
        Me.btn_expandir.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn_expandir.Location = New System.Drawing.Point(921, 1)
        Me.btn_expandir.Margin = New System.Windows.Forms.Padding(6)
        Me.btn_expandir.Name = "btn_expandir"
        Me.btn_expandir.Size = New System.Drawing.Size(43, 39)
        Me.btn_expandir.TabIndex = 90
        Me.btn_expandir.UseVisualStyleBackColor = False
        '
        'btn_fechar
        '
        Me.btn_fechar.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_fechar.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn_fechar.BackgroundImage = Global.Ordem_de_Servico.My.Resources.Resources._5
        Me.btn_fechar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btn_fechar.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(49, Byte), Integer), CType(CType(49, Byte), Integer))
        Me.btn_fechar.FlatAppearance.BorderSize = 0
        Me.btn_fechar.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(49, Byte), Integer), CType(CType(49, Byte), Integer))
        Me.btn_fechar.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(49, Byte), Integer), CType(CType(49, Byte), Integer))
        Me.btn_fechar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_fechar.Font = New System.Drawing.Font("Arial Narrow", 16.0!, System.Drawing.FontStyle.Bold)
        Me.btn_fechar.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn_fechar.Location = New System.Drawing.Point(964, 1)
        Me.btn_fechar.Margin = New System.Windows.Forms.Padding(6)
        Me.btn_fechar.Name = "btn_fechar"
        Me.btn_fechar.Size = New System.Drawing.Size(41, 39)
        Me.btn_fechar.TabIndex = 89
        Me.btn_fechar.UseVisualStyleBackColor = False
        '
        'btn_home
        '
        Me.btn_home.BackColor = System.Drawing.Color.OrangeRed
        Me.btn_home.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btn_home.FlatAppearance.BorderColor = System.Drawing.Color.OrangeRed
        Me.btn_home.FlatAppearance.BorderSize = 0
        Me.btn_home.FlatAppearance.MouseDownBackColor = System.Drawing.Color.OrangeRed
        Me.btn_home.FlatAppearance.MouseOverBackColor = System.Drawing.Color.OrangeRed
        Me.btn_home.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_home.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_home.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn_home.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_home.Location = New System.Drawing.Point(163, 1)
        Me.btn_home.Margin = New System.Windows.Forms.Padding(6)
        Me.btn_home.Name = "btn_home"
        Me.btn_home.Size = New System.Drawing.Size(82, 51)
        Me.btn_home.TabIndex = 81
        Me.btn_home.Text = "Home"
        Me.btn_home.UseVisualStyleBackColor = False
        '
        'btn_os
        '
        Me.btn_os.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.btn_os.BackgroundImage = Global.Ordem_de_Servico.My.Resources.Resources.Design_sem_nome__7_
        Me.btn_os.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.btn_os.FlatAppearance.BorderSize = 0
        Me.btn_os.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.btn_os.FlatAppearance.MouseOverBackColor = System.Drawing.Color.OrangeRed
        Me.btn_os.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_os.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Bold)
        Me.btn_os.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn_os.Location = New System.Drawing.Point(257, 1)
        Me.btn_os.Margin = New System.Windows.Forms.Padding(6)
        Me.btn_os.Name = "btn_os"
        Me.btn_os.Size = New System.Drawing.Size(141, 51)
        Me.btn_os.TabIndex = 1
        Me.btn_os.Text = "Solicitar OS"
        Me.btn_os.UseVisualStyleBackColor = False
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AllowUserToOrderColumns = True
        Me.DataGridView1.AllowUserToResizeColumns = False
        Me.DataGridView1.AllowUserToResizeRows = False
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.BackgroundColor = System.Drawing.SystemColors.ControlLightLight
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView1.DefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridView1.GridColor = System.Drawing.SystemColors.ButtonFace
        Me.DataGridView1.Location = New System.Drawing.Point(50, 154)
        Me.DataGridView1.MultiSelect = False
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(920, 380)
        Me.DataGridView1.TabIndex = 84
        '
        'txt_situacao
        '
        Me.txt_situacao.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_situacao.FormattingEnabled = True
        Me.txt_situacao.Location = New System.Drawing.Point(298, 117)
        Me.txt_situacao.Margin = New System.Windows.Forms.Padding(5)
        Me.txt_situacao.Name = "txt_situacao"
        Me.txt_situacao.Size = New System.Drawing.Size(182, 30)
        Me.txt_situacao.TabIndex = 86
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial", 12.0!)
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label1.Location = New System.Drawing.Point(47, 93)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(79, 18)
        Me.Label1.TabIndex = 87
        Me.Label1.Text = "Pesquisar"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Arial", 12.0!)
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label2.Location = New System.Drawing.Point(295, 93)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(70, 18)
        Me.Label2.TabIndex = 88
        Me.Label2.Text = "Situação"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.OrangeRed
        Me.Button1.BackgroundImage = Global.Ordem_de_Servico.My.Resources.Resources.icon_f5
        Me.Button1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.Button1.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(49, Byte), Integer), CType(CType(49, Byte), Integer))
        Me.Button1.FlatAppearance.BorderSize = 0
        Me.Button1.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(49, Byte), Integer), CType(CType(49, Byte), Integer))
        Me.Button1.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(49, Byte), Integer), CType(CType(49, Byte), Integer))
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Arial Narrow", 16.0!, System.Drawing.FontStyle.Bold)
        Me.Button1.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Button1.Location = New System.Drawing.Point(491, 117)
        Me.Button1.Margin = New System.Windows.Forms.Padding(6)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(34, 30)
        Me.Button1.TabIndex = 85
        Me.Button1.UseVisualStyleBackColor = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.PictureBox1.Image = Global.Ordem_de_Servico.My.Resources.Resources.logo
        Me.PictureBox1.Location = New System.Drawing.Point(415, 547)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(192, 23)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 92
        Me.PictureBox1.TabStop = False
        '
        'Tela_Inicial
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.ClientSize = New System.Drawing.Size(1019, 622)
        Me.ControlBox = False
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txt_situacao)
        Me.Controls.Add(Me.txt_fer)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.DataGridView1)
        Me.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.HelpButton = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Tela_Inicial"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " "
        Me.Panel1.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txt_fer As TextBox
    Friend WithEvents btn_relatorio As Button
    Friend WithEvents btn_servico As Button
    Friend WithEvents btn_maq As Button
    Friend WithEvents btn_os As Button
    Friend WithEvents btn_home As Button
    Friend WithEvents btn_user As Button
    Friend WithEvents Panel1 As Panel
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents txt_situacao As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents btn_minimizar As Button
    Friend WithEvents btn_expandir As Button
    Friend WithEvents btn_fechar As Button
    Friend WithEvents PictureBox1 As PictureBox
End Class
