<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Relatorio
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Relatorio))
        Me.txt_fer = New System.Windows.Forms.TextBox()
        Me.Lista_OS = New System.Windows.Forms.ListView()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txt_data = New System.Windows.Forms.MaskedTextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txt_os = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txt_op = New System.Windows.Forms.ComboBox()
        Me.cbo_maq = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.btn_limpar = New System.Windows.Forms.Button()
        Me.txt_data_fim = New System.Windows.Forms.MaskedTextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.btn_servico = New System.Windows.Forms.Button()
        Me.btn_relatorio = New System.Windows.Forms.Button()
        Me.btn_maq = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btn_minimizar = New System.Windows.Forms.Button()
        Me.btn_home = New System.Windows.Forms.Button()
        Me.btn_expandir = New System.Windows.Forms.Button()
        Me.btn_fechar = New System.Windows.Forms.Button()
        Me.btn_os = New System.Windows.Forms.Button()
        Me.btn_user = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.btn_excel = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Lista_Parada = New System.Windows.Forms.ListView()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txt_fer
        '
        Me.txt_fer.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt_fer.Font = New System.Drawing.Font("Arial Narrow", 14.0!)
        Me.txt_fer.Location = New System.Drawing.Point(114, 121)
        Me.txt_fer.Margin = New System.Windows.Forms.Padding(6)
        Me.txt_fer.Name = "txt_fer"
        Me.txt_fer.Size = New System.Drawing.Size(131, 29)
        Me.txt_fer.TabIndex = 2
        '
        'Lista_OS
        '
        Me.Lista_OS.AllowColumnReorder = True
        Me.Lista_OS.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lista_OS.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Lista_OS.Font = New System.Drawing.Font("Montserrat Medium", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lista_OS.ForeColor = System.Drawing.Color.FromArgb(CType(CType(32, Byte), Integer), CType(CType(32, Byte), Integer), CType(CType(32, Byte), Integer))
        Me.Lista_OS.FullRowSelect = True
        Me.Lista_OS.GridLines = True
        Me.Lista_OS.Location = New System.Drawing.Point(292, 121)
        Me.Lista_OS.MultiSelect = False
        Me.Lista_OS.Name = "Lista_OS"
        Me.Lista_OS.Size = New System.Drawing.Size(753, 211)
        Me.Lista_OS.TabIndex = 69
        Me.Lista_OS.UseCompatibleStateImageBehavior = False
        Me.Lista_OS.View = System.Windows.Forms.View.Details
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label2.Font = New System.Drawing.Font("Arial", 12.0!)
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label2.Location = New System.Drawing.Point(111, 97)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(89, 18)
        Me.Label2.TabIndex = 73
        Me.Label2.Text = "Ferramenta"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label3.Font = New System.Drawing.Font("Arial", 12.0!)
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label3.Location = New System.Drawing.Point(44, 282)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(81, 18)
        Me.Label3.TabIndex = 74
        Me.Label3.Text = "Data Início"
        '
        'txt_data
        '
        Me.txt_data.Font = New System.Drawing.Font("Arial Narrow", 14.0!)
        Me.txt_data.Location = New System.Drawing.Point(47, 303)
        Me.txt_data.Mask = "00/00/0000"
        Me.txt_data.Name = "txt_data"
        Me.txt_data.Size = New System.Drawing.Size(93, 29)
        Me.txt_data.TabIndex = 3
        Me.txt_data.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txt_data.ValidatingType = GetType(Date)
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label5.Font = New System.Drawing.Font("Arial", 12.0!)
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label5.Location = New System.Drawing.Point(47, 97)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(31, 18)
        Me.Label5.TabIndex = 81
        Me.Label5.Text = "OS"
        '
        'txt_os
        '
        Me.txt_os.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txt_os.Font = New System.Drawing.Font("Arial Narrow", 14.0!)
        Me.txt_os.Location = New System.Drawing.Point(47, 121)
        Me.txt_os.Margin = New System.Windows.Forms.Padding(6)
        Me.txt_os.Name = "txt_os"
        Me.txt_os.Size = New System.Drawing.Size(55, 29)
        Me.txt_os.TabIndex = 1
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label7.Font = New System.Drawing.Font("Arial", 12.0!)
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label7.Location = New System.Drawing.Point(43, 157)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(78, 18)
        Me.Label7.TabIndex = 85
        Me.Label7.Text = "Operação"
        '
        'txt_op
        '
        Me.txt_op.Font = New System.Drawing.Font("Arial Narrow", 14.0!)
        Me.txt_op.FormattingEnabled = True
        Me.txt_op.Location = New System.Drawing.Point(47, 180)
        Me.txt_op.Name = "txt_op"
        Me.txt_op.Size = New System.Drawing.Size(198, 31)
        Me.txt_op.TabIndex = 6
        '
        'cbo_maq
        '
        Me.cbo_maq.Font = New System.Drawing.Font("Arial Narrow", 14.0!)
        Me.cbo_maq.FormattingEnabled = True
        Me.cbo_maq.Location = New System.Drawing.Point(47, 243)
        Me.cbo_maq.Name = "cbo_maq"
        Me.cbo_maq.Size = New System.Drawing.Size(198, 31)
        Me.cbo_maq.TabIndex = 7
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label8.Font = New System.Drawing.Font("Arial", 12.0!)
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label8.Location = New System.Drawing.Point(44, 220)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(68, 18)
        Me.Label8.TabIndex = 87
        Me.Label8.Text = "Máquina"
        '
        'btn_limpar
        '
        Me.btn_limpar.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btn_limpar.BackColor = System.Drawing.Color.Transparent
        Me.btn_limpar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btn_limpar.FlatAppearance.BorderColor = System.Drawing.Color.DarkOrange
        Me.btn_limpar.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.btn_limpar.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange
        Me.btn_limpar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_limpar.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Bold)
        Me.btn_limpar.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn_limpar.Location = New System.Drawing.Point(47, 510)
        Me.btn_limpar.Name = "btn_limpar"
        Me.btn_limpar.Size = New System.Drawing.Size(93, 39)
        Me.btn_limpar.TabIndex = 10
        Me.btn_limpar.Text = "RESETAR"
        Me.btn_limpar.UseVisualStyleBackColor = False
        '
        'txt_data_fim
        '
        Me.txt_data_fim.Font = New System.Drawing.Font("Arial Narrow", 14.0!)
        Me.txt_data_fim.Location = New System.Drawing.Point(152, 303)
        Me.txt_data_fim.Mask = "00/00/0000"
        Me.txt_data_fim.Name = "txt_data_fim"
        Me.txt_data_fim.Size = New System.Drawing.Size(93, 29)
        Me.txt_data_fim.TabIndex = 4
        Me.txt_data_fim.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txt_data_fim.ValidatingType = GetType(Date)
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label9.Font = New System.Drawing.Font("Arial", 12.0!)
        Me.Label9.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label9.Location = New System.Drawing.Point(149, 282)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(73, 18)
        Me.Label9.TabIndex = 90
        Me.Label9.Text = "Data Fim"
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
        Me.btn_servico.TabIndex = 93
        Me.btn_servico.Text = "Apontar Serviço"
        Me.btn_servico.UseVisualStyleBackColor = False
        '
        'btn_relatorio
        '
        Me.btn_relatorio.BackColor = System.Drawing.Color.OrangeRed
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
        Me.btn_relatorio.TabIndex = 94
        Me.btn_relatorio.Text = "Relatório"
        Me.btn_relatorio.UseVisualStyleBackColor = False
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
        Me.btn_maq.TabIndex = 95
        Me.btn_maq.Text = " Máquina"
        Me.btn_maq.UseVisualStyleBackColor = False
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Panel1.Controls.Add(Me.btn_minimizar)
        Me.Panel1.Controls.Add(Me.btn_home)
        Me.Panel1.Controls.Add(Me.btn_expandir)
        Me.Panel1.Controls.Add(Me.btn_maq)
        Me.Panel1.Controls.Add(Me.btn_fechar)
        Me.Panel1.Controls.Add(Me.btn_relatorio)
        Me.Panel1.Controls.Add(Me.btn_os)
        Me.Panel1.Controls.Add(Me.btn_servico)
        Me.Panel1.Controls.Add(Me.btn_user)
        Me.Panel1.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Panel1.Location = New System.Drawing.Point(-1, -1)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1095, 76)
        Me.Panel1.TabIndex = 97
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
        Me.btn_minimizar.Location = New System.Drawing.Point(951, 0)
        Me.btn_minimizar.Margin = New System.Windows.Forms.Padding(6)
        Me.btn_minimizar.Name = "btn_minimizar"
        Me.btn_minimizar.Size = New System.Drawing.Size(42, 39)
        Me.btn_minimizar.TabIndex = 103
        Me.btn_minimizar.UseVisualStyleBackColor = False
        '
        'btn_home
        '
        Me.btn_home.BackColor = System.Drawing.Color.Black
        Me.btn_home.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btn_home.FlatAppearance.BorderColor = System.Drawing.Color.OrangeRed
        Me.btn_home.FlatAppearance.BorderSize = 0
        Me.btn_home.FlatAppearance.MouseDownBackColor = System.Drawing.Color.OrangeRed
        Me.btn_home.FlatAppearance.MouseOverBackColor = System.Drawing.Color.OrangeRed
        Me.btn_home.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_home.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Bold)
        Me.btn_home.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn_home.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_home.Location = New System.Drawing.Point(163, 1)
        Me.btn_home.Margin = New System.Windows.Forms.Padding(6)
        Me.btn_home.Name = "btn_home"
        Me.btn_home.Size = New System.Drawing.Size(83, 51)
        Me.btn_home.TabIndex = 98
        Me.btn_home.Text = "Home"
        Me.btn_home.UseVisualStyleBackColor = False
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
        Me.btn_expandir.Location = New System.Drawing.Point(993, 0)
        Me.btn_expandir.Margin = New System.Windows.Forms.Padding(6)
        Me.btn_expandir.Name = "btn_expandir"
        Me.btn_expandir.Size = New System.Drawing.Size(43, 39)
        Me.btn_expandir.TabIndex = 102
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
        Me.btn_fechar.Location = New System.Drawing.Point(1036, 0)
        Me.btn_fechar.Margin = New System.Windows.Forms.Padding(6)
        Me.btn_fechar.Name = "btn_fechar"
        Me.btn_fechar.Size = New System.Drawing.Size(41, 39)
        Me.btn_fechar.TabIndex = 100
        Me.btn_fechar.UseVisualStyleBackColor = False
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
        Me.btn_os.TabIndex = 92
        Me.btn_os.Text = "Solicitar OS"
        Me.btn_os.UseVisualStyleBackColor = False
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
        Me.btn_user.TabIndex = 98
        Me.btn_user.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label1.Font = New System.Drawing.Font("Arial", 12.0!)
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label1.Location = New System.Drawing.Point(289, 97)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(86, 18)
        Me.Label1.TabIndex = 99
        Me.Label1.Text = "Operações"
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.PictureBox1.Image = Global.Ordem_de_Servico.My.Resources.Resources.logo
        Me.PictureBox1.Location = New System.Drawing.Point(451, 568)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(192, 23)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 91
        Me.PictureBox1.TabStop = False
        '
        'btn_excel
        '
        Me.btn_excel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btn_excel.BackColor = System.Drawing.Color.Transparent
        Me.btn_excel.BackgroundImage = Global.Ordem_de_Servico.My.Resources.Resources.excel
        Me.btn_excel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btn_excel.FlatAppearance.BorderColor = System.Drawing.Color.Green
        Me.btn_excel.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.btn_excel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Green
        Me.btn_excel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_excel.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_excel.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn_excel.Location = New System.Drawing.Point(152, 510)
        Me.btn_excel.Name = "btn_excel"
        Me.btn_excel.Size = New System.Drawing.Size(93, 39)
        Me.btn_excel.TabIndex = 11
        Me.btn_excel.UseVisualStyleBackColor = False
        '
        'Label4
        '
        Me.Label4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label4.Font = New System.Drawing.Font("Arial", 12.0!)
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label4.Location = New System.Drawing.Point(289, 338)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(68, 18)
        Me.Label4.TabIndex = 101
        Me.Label4.Text = "Paradas"
        '
        'Lista_Parada
        '
        Me.Lista_Parada.AllowColumnReorder = True
        Me.Lista_Parada.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lista_Parada.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Lista_Parada.Font = New System.Drawing.Font("Montserrat Medium", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lista_Parada.ForeColor = System.Drawing.Color.FromArgb(CType(CType(32, Byte), Integer), CType(CType(32, Byte), Integer), CType(CType(32, Byte), Integer))
        Me.Lista_Parada.FullRowSelect = True
        Me.Lista_Parada.GridLines = True
        Me.Lista_Parada.Location = New System.Drawing.Point(292, 359)
        Me.Lista_Parada.MultiSelect = False
        Me.Lista_Parada.Name = "Lista_Parada"
        Me.Lista_Parada.Size = New System.Drawing.Size(753, 190)
        Me.Lista_Parada.TabIndex = 100
        Me.Lista_Parada.UseCompatibleStateImageBehavior = False
        Me.Lista_Parada.View = System.Windows.Forms.View.Details
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AllowUserToOrderColumns = True
        Me.DataGridView1.AllowUserToResizeColumns = False
        Me.DataGridView1.AllowUserToResizeRows = False
        Me.DataGridView1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
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
        Me.DataGridView1.Location = New System.Drawing.Point(46, 347)
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
        Me.DataGridView1.Size = New System.Drawing.Size(199, 150)
        Me.DataGridView1.TabIndex = 102
        '
        'Relatorio
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.ClientSize = New System.Drawing.Size(1089, 643)
        Me.ControlBox = False
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Lista_Parada)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.txt_data_fim)
        Me.Controls.Add(Me.cbo_maq)
        Me.Controls.Add(Me.txt_op)
        Me.Controls.Add(Me.txt_os)
        Me.Controls.Add(Me.btn_excel)
        Me.Controls.Add(Me.txt_data)
        Me.Controls.Add(Me.txt_fer)
        Me.Controls.Add(Me.Lista_OS)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btn_limpar)
        Me.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.HelpButton = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Relatorio"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " "
        Me.Panel1.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txt_fer As TextBox
    Friend WithEvents Lista_OS As ListView
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents txt_data As MaskedTextBox
    Friend WithEvents btn_excel As Button
    Friend WithEvents Label5 As Label
    Friend WithEvents txt_os As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents txt_op As ComboBox
    Friend WithEvents cbo_maq As ComboBox
    Friend WithEvents Label8 As Label
    Friend WithEvents btn_limpar As Button
    Friend WithEvents txt_data_fim As MaskedTextBox
    Friend WithEvents Label9 As Label
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents btn_servico As Button
    Friend WithEvents btn_os As Button
    Friend WithEvents btn_relatorio As Button
    Friend WithEvents btn_maq As Button
    Friend WithEvents Panel1 As Panel
    Friend WithEvents btn_user As Button
    Friend WithEvents btn_home As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Label10 As Label
    Friend WithEvents btn_minimizar As Button
    Friend WithEvents btn_expandir As Button
    Friend WithEvents btn_fechar As Button
    Friend WithEvents Label4 As Label
    Friend WithEvents Lista_Parada As ListView
    Friend WithEvents DataGridView1 As DataGridView
End Class
