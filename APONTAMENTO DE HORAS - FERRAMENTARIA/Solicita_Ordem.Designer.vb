<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Solicita_Ordem
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Solicita_Ordem))
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btn_minimizar = New System.Windows.Forms.Button()
        Me.btn_home = New System.Windows.Forms.Button()
        Me.btn_expandir = New System.Windows.Forms.Button()
        Me.btn_fechar = New System.Windows.Forms.Button()
        Me.btn_os = New System.Windows.Forms.Button()
        Me.btn_user = New System.Windows.Forms.Button()
        Me.txt_OS = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.txt_secao = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.txt_situ = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txt_sub_conta = New System.Windows.Forms.ComboBox()
        Me.txt_conta = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txt_desc = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txt_pos = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txt_fer = New System.Windows.Forms.TextBox()
        Me.btn_salvar = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txt_limpar = New System.Windows.Forms.Button()
        Me.Lista_Request = New System.Windows.Forms.ListView()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txt_motivo = New System.Windows.Forms.ComboBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.txt_tp = New System.Windows.Forms.ComboBox()
        Me.txt_qtd = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Panel3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel3
        '
        Me.Panel3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel3.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Panel3.Controls.Add(Me.btn_minimizar)
        Me.Panel3.Controls.Add(Me.btn_home)
        Me.Panel3.Controls.Add(Me.btn_expandir)
        Me.Panel3.Controls.Add(Me.btn_fechar)
        Me.Panel3.Controls.Add(Me.btn_os)
        Me.Panel3.Controls.Add(Me.btn_user)
        Me.Panel3.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Panel3.Location = New System.Drawing.Point(-1, -1)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(1023, 76)
        Me.Panel3.TabIndex = 107
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
        Me.btn_minimizar.Location = New System.Drawing.Point(827, 0)
        Me.btn_minimizar.Margin = New System.Windows.Forms.Padding(6)
        Me.btn_minimizar.Name = "btn_minimizar"
        Me.btn_minimizar.Size = New System.Drawing.Size(42, 39)
        Me.btn_minimizar.TabIndex = 121
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
        Me.btn_home.Size = New System.Drawing.Size(102, 51)
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
        Me.btn_expandir.Location = New System.Drawing.Point(869, 0)
        Me.btn_expandir.Margin = New System.Windows.Forms.Padding(6)
        Me.btn_expandir.Name = "btn_expandir"
        Me.btn_expandir.Size = New System.Drawing.Size(43, 39)
        Me.btn_expandir.TabIndex = 120
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
        Me.btn_fechar.Location = New System.Drawing.Point(912, 0)
        Me.btn_fechar.Margin = New System.Windows.Forms.Padding(6)
        Me.btn_fechar.Name = "btn_fechar"
        Me.btn_fechar.Size = New System.Drawing.Size(41, 39)
        Me.btn_fechar.TabIndex = 118
        Me.btn_fechar.UseVisualStyleBackColor = False
        '
        'btn_os
        '
        Me.btn_os.BackColor = System.Drawing.Color.OrangeRed
        Me.btn_os.BackgroundImage = Global.Ordem_de_Servico.My.Resources.Resources.Design_sem_nome__7_
        Me.btn_os.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.btn_os.FlatAppearance.BorderSize = 0
        Me.btn_os.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.btn_os.FlatAppearance.MouseOverBackColor = System.Drawing.Color.OrangeRed
        Me.btn_os.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_os.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Bold)
        Me.btn_os.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn_os.Location = New System.Drawing.Point(277, 1)
        Me.btn_os.Margin = New System.Windows.Forms.Padding(6)
        Me.btn_os.Name = "btn_os"
        Me.btn_os.Size = New System.Drawing.Size(163, 51)
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
        'txt_OS
        '
        Me.txt_OS.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.txt_OS.Enabled = False
        Me.txt_OS.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_OS.Location = New System.Drawing.Point(8, 23)
        Me.txt_OS.Margin = New System.Windows.Forms.Padding(5)
        Me.txt_OS.Name = "txt_OS"
        Me.txt_OS.Size = New System.Drawing.Size(91, 29)
        Me.txt_OS.TabIndex = 1
        '
        'Label18
        '
        Me.Label18.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label18.AutoSize = True
        Me.Label18.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.Location = New System.Drawing.Point(314, 62)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(60, 18)
        Me.Label18.TabIndex = 52
        Me.Label18.Text = "Seção*"
        '
        'txt_secao
        '
        Me.txt_secao.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.txt_secao.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_secao.Location = New System.Drawing.Point(317, 82)
        Me.txt_secao.Margin = New System.Windows.Forms.Padding(5)
        Me.txt_secao.Name = "txt_secao"
        Me.txt_secao.Size = New System.Drawing.Size(135, 29)
        Me.txt_secao.TabIndex = 16
        '
        'Label16
        '
        Me.Label16.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(204, 60)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(85, 18)
        Me.Label16.TabIndex = 48
        Me.Label16.Text = "Sub-conta*"
        '
        'Label15
        '
        Me.Label15.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(104, 60)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(56, 18)
        Me.Label15.TabIndex = 46
        Me.Label15.Text = "Conta*"
        '
        'txt_situ
        '
        Me.txt_situ.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.txt_situ.Enabled = False
        Me.txt_situ.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_situ.Location = New System.Drawing.Point(759, 23)
        Me.txt_situ.Margin = New System.Windows.Forms.Padding(5)
        Me.txt_situ.Name = "txt_situ"
        Me.txt_situ.Size = New System.Drawing.Size(150, 29)
        Me.txt_situ.TabIndex = 5
        Me.txt_situ.Text = "ENVIADA"
        '
        'Label4
        '
        Me.Label4.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(755, 2)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(70, 18)
        Me.Label4.TabIndex = 91
        Me.Label4.Text = "Situação"
        '
        'txt_sub_conta
        '
        Me.txt_sub_conta.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.txt_sub_conta.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_sub_conta.FormattingEnabled = True
        Me.txt_sub_conta.Location = New System.Drawing.Point(208, 80)
        Me.txt_sub_conta.Margin = New System.Windows.Forms.Padding(5)
        Me.txt_sub_conta.Name = "txt_sub_conta"
        Me.txt_sub_conta.Size = New System.Drawing.Size(99, 30)
        Me.txt_sub_conta.TabIndex = 15
        '
        'txt_conta
        '
        Me.txt_conta.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.txt_conta.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_conta.FormattingEnabled = True
        Me.txt_conta.Location = New System.Drawing.Point(108, 80)
        Me.txt_conta.Margin = New System.Windows.Forms.Padding(5)
        Me.txt_conta.Name = "txt_conta"
        Me.txt_conta.Size = New System.Drawing.Size(90, 30)
        Me.txt_conta.TabIndex = 14
        '
        'Label7
        '
        Me.Label7.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(523, 2)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(80, 18)
        Me.Label7.TabIndex = 28
        Me.Label7.Text = "Descrição"
        '
        'txt_desc
        '
        Me.txt_desc.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.txt_desc.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_desc.Location = New System.Drawing.Point(527, 23)
        Me.txt_desc.Margin = New System.Windows.Forms.Padding(5)
        Me.txt_desc.Name = "txt_desc"
        Me.txt_desc.Size = New System.Drawing.Size(222, 29)
        Me.txt_desc.TabIndex = 4
        '
        'Label6
        '
        Me.Label6.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(313, 2)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(72, 18)
        Me.Label6.TabIndex = 26
        Me.Label6.Text = "Posição*"
        '
        'txt_pos
        '
        Me.txt_pos.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.txt_pos.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_pos.Location = New System.Drawing.Point(317, 23)
        Me.txt_pos.Margin = New System.Windows.Forms.Padding(5)
        Me.txt_pos.Name = "txt_pos"
        Me.txt_pos.Size = New System.Drawing.Size(200, 29)
        Me.txt_pos.TabIndex = 3
        '
        'Label5
        '
        Me.Label5.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(105, 2)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(95, 18)
        Me.Label5.TabIndex = 24
        Me.Label5.Text = "Ferramenta*"
        '
        'txt_fer
        '
        Me.txt_fer.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.txt_fer.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_fer.Location = New System.Drawing.Point(107, 23)
        Me.txt_fer.Margin = New System.Windows.Forms.Padding(5)
        Me.txt_fer.Name = "txt_fer"
        Me.txt_fer.Size = New System.Drawing.Size(200, 29)
        Me.txt_fer.TabIndex = 2
        '
        'btn_salvar
        '
        Me.btn_salvar.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btn_salvar.BackColor = System.Drawing.Color.Transparent
        Me.btn_salvar.FlatAppearance.BorderColor = System.Drawing.Color.OrangeRed
        Me.btn_salvar.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.btn_salvar.FlatAppearance.MouseOverBackColor = System.Drawing.Color.OrangeRed
        Me.btn_salvar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_salvar.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_salvar.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.btn_salvar.Location = New System.Drawing.Point(758, 81)
        Me.btn_salvar.Name = "btn_salvar"
        Me.btn_salvar.Size = New System.Drawing.Size(74, 30)
        Me.btn_salvar.TabIndex = 18
        Me.btn_salvar.Text = "ENVIAR"
        Me.btn_salvar.UseVisualStyleBackColor = False
        '
        'Label2
        '
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(4, 2)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(31, 18)
        Me.Label2.TabIndex = 80
        Me.Label2.Text = "OS"
        '
        'txt_limpar
        '
        Me.txt_limpar.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.txt_limpar.BackColor = System.Drawing.Color.Transparent
        Me.txt_limpar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.txt_limpar.FlatAppearance.BorderColor = System.Drawing.Color.DarkOrange
        Me.txt_limpar.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.txt_limpar.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DarkOrange
        Me.txt_limpar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.txt_limpar.Font = New System.Drawing.Font("Arial Narrow", 12.0!, System.Drawing.FontStyle.Bold)
        Me.txt_limpar.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.txt_limpar.Location = New System.Drawing.Point(838, 81)
        Me.txt_limpar.Name = "txt_limpar"
        Me.txt_limpar.Size = New System.Drawing.Size(71, 30)
        Me.txt_limpar.TabIndex = 107
        Me.txt_limpar.Text = "LIMPAR"
        Me.txt_limpar.UseVisualStyleBackColor = False
        '
        'Lista_Request
        '
        Me.Lista_Request.AllowColumnReorder = True
        Me.Lista_Request.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Lista_Request.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Lista_Request.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lista_Request.ForeColor = System.Drawing.Color.FromArgb(CType(CType(32, Byte), Integer), CType(CType(32, Byte), Integer), CType(CType(32, Byte), Integer))
        Me.Lista_Request.FullRowSelect = True
        Me.Lista_Request.GridLines = True
        Me.Lista_Request.Location = New System.Drawing.Point(7, 119)
        Me.Lista_Request.MultiSelect = False
        Me.Lista_Request.Name = "Lista_Request"
        Me.Lista_Request.Size = New System.Drawing.Size(902, 337)
        Me.Lista_Request.TabIndex = 19
        Me.Lista_Request.UseCompatibleStateImageBehavior = False
        Me.Lista_Request.View = System.Windows.Forms.View.Details
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.txt_motivo)
        Me.Panel2.Controls.Add(Me.Label13)
        Me.Panel2.Controls.Add(Me.txt_tp)
        Me.Panel2.Controls.Add(Me.Lista_Request)
        Me.Panel2.Controls.Add(Me.txt_limpar)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.btn_salvar)
        Me.Panel2.Controls.Add(Me.txt_fer)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.txt_pos)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Controls.Add(Me.txt_desc)
        Me.Panel2.Controls.Add(Me.Label7)
        Me.Panel2.Controls.Add(Me.txt_conta)
        Me.Panel2.Controls.Add(Me.txt_sub_conta)
        Me.Panel2.Controls.Add(Me.txt_qtd)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.Label14)
        Me.Panel2.Controls.Add(Me.txt_situ)
        Me.Panel2.Controls.Add(Me.Label15)
        Me.Panel2.Controls.Add(Me.Label16)
        Me.Panel2.Controls.Add(Me.txt_secao)
        Me.Panel2.Controls.Add(Me.Label18)
        Me.Panel2.Controls.Add(Me.txt_OS)
        Me.Panel2.Location = New System.Drawing.Point(43, 86)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(915, 465)
        Me.Panel2.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(592, 59)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(60, 18)
        Me.Label1.TabIndex = 111
        Me.Label1.Text = "Motivo*"
        '
        'txt_motivo
        '
        Me.txt_motivo.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.txt_motivo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.txt_motivo.Enabled = False
        Me.txt_motivo.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_motivo.FormattingEnabled = True
        Me.txt_motivo.Location = New System.Drawing.Point(596, 80)
        Me.txt_motivo.Margin = New System.Windows.Forms.Padding(5)
        Me.txt_motivo.Name = "txt_motivo"
        Me.txt_motivo.Size = New System.Drawing.Size(153, 30)
        Me.txt_motivo.TabIndex = 110
        '
        'Label13
        '
        Me.Label13.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(458, 60)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(44, 18)
        Me.Label13.TabIndex = 109
        Me.Label13.Text = "Tipo*"
        '
        'txt_tp
        '
        Me.txt_tp.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.txt_tp.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.txt_tp.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_tp.FormattingEnabled = True
        Me.txt_tp.Location = New System.Drawing.Point(462, 81)
        Me.txt_tp.Margin = New System.Windows.Forms.Padding(5)
        Me.txt_tp.Name = "txt_tp"
        Me.txt_tp.Size = New System.Drawing.Size(124, 30)
        Me.txt_tp.TabIndex = 108
        '
        'txt_qtd
        '
        Me.txt_qtd.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.txt_qtd.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_qtd.Location = New System.Drawing.Point(9, 81)
        Me.txt_qtd.Margin = New System.Windows.Forms.Padding(5)
        Me.txt_qtd.Name = "txt_qtd"
        Me.txt_qtd.Size = New System.Drawing.Size(90, 29)
        Me.txt_qtd.TabIndex = 12
        '
        'Label14
        '
        Me.Label14.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(5, 61)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(43, 18)
        Me.Label14.TabIndex = 42
        Me.Label14.Text = "Qtd.*"
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.PictureBox1.Image = Global.Ordem_de_Servico.My.Resources.Resources.logo
        Me.PictureBox1.Location = New System.Drawing.Point(406, 558)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(192, 23)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 122
        Me.PictureBox1.TabStop = False
        '
        'Solicita_Ordem
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Black
        Me.ClientSize = New System.Drawing.Size(1008, 632)
        Me.ControlBox = False
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel2)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Solicita_Ordem"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " "
        Me.Panel3.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel3 As Panel
    Friend WithEvents btn_home As Button
    Friend WithEvents btn_os As Button
    Friend WithEvents btn_user As Button
    Friend WithEvents txt_OS As TextBox
    Friend WithEvents Label18 As Label
    Friend WithEvents txt_secao As TextBox
    Friend WithEvents Label16 As Label
    Friend WithEvents Label15 As Label
    Friend WithEvents txt_situ As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents txt_sub_conta As ComboBox
    Friend WithEvents txt_conta As ComboBox
    Friend WithEvents Label7 As Label
    Friend WithEvents txt_desc As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents txt_pos As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents txt_fer As TextBox
    Friend WithEvents btn_salvar As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents txt_limpar As Button
    Friend WithEvents Lista_Request As ListView
    Friend WithEvents Panel2 As Panel
    Friend WithEvents txt_qtd As TextBox
    Friend WithEvents Label14 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents txt_motivo As ComboBox
    Friend WithEvents Label13 As Label
    Friend WithEvents txt_tp As ComboBox
    Friend WithEvents btn_minimizar As Button
    Friend WithEvents btn_expandir As Button
    Friend WithEvents btn_fechar As Button
    Friend WithEvents PictureBox1 As PictureBox
End Class
