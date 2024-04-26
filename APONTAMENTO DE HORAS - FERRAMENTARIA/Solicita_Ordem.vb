

Imports System.IO

Public Class Solicita_Ordem
    Public COD_FORM As Integer
    Public REGISTRO As String

    Private Sub Ordem_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        permissoes()
        Preencher_cbo()

        If id_OS = 0 Then
            conectar()
            comandoSQL.CommandText = " SELECT TOP 1 OS_ID FROM TAB_OS_RELATORIO ORDER BY OS_ID DESC"
            objDataReader = comandoSQL.ExecuteReader
            If objDataReader.Read Then
                COD_FORM = objDataReader(0) + 1
            Else
                COD_FORM = 1
            End If
            objDataReader.Close()
            desconectar()
            txt_OS.Text = COD_FORM
        Else
            COD_FORM = txt_OS.Text
        End If
        Lista_Req()
    End Sub

    Private Sub btn_salvar_Click(sender As Object, e As EventArgs) Handles btn_salvar.Click
        If txt_fer.Text = "" Or txt_pos.Text = "" Or txt_secao.Text = "" Or txt_secao.Text = "" Or txt_qtd.Text = "" Or txt_conta.Text = "" Or txt_sub_conta.Text = "" Then
            MessageBox.Show("Preencha os campos obrigatórios")
            Return
        End If
        Try
            conectar()
            objDataReader.Close()
            comandoSQL.CommandText = "INSERT INTO TAB_OS_RELATORIO (OS_ID, OS_FERRAMENTA, OS_POSICAO, OS_DESCRICAO, OS_QUANTIDADE, OS_TIPO, OS_MOTIVO, "
            comandoSQL.CommandText &= "OS_CONTA,  OS_SUB_CONTA, OS_SECAO, OS_REG_RESPONSAVEL, OS_SITUACAO, OS_DATA_ENVIADA) VALUES ('" & COD_FORM & "', '" & txt_fer.Text & "', '" & txt_pos.Text & "', '" & txt_desc.Text & "', "
            comandoSQL.CommandText &= "'" & txt_qtd.Text & "', '" & txt_tp.Text & "', '" & txt_motivo.Text & "', '" & txt_conta.Text & "', '" & txt_sub_conta.Text & "', '" & txt_secao.Text & "', '" & REGISTRO & "', '" & txt_situ.Text & "', '" & DateTime.Now & "')"
            comandoSQL.ExecuteReader()
            desconectar()
            Resetar()
            MessageBox.Show("Ordem solicitada.")
        Catch ex As Exception
            MessageBox.Show("Ocorreu um erro: " & ex.Message)
        End Try
        Resetar()
        Lista_Req()
    End Sub

    Sub Preencher_cbo()
        txt_tp.Items.Add("")
        txt_tp.Items.Add("RETRABALHO")
        txt_tp.Items.Add("PEÇA NOVA")

        txt_conta.Items.Add("7000")
        txt_conta.Items.Add("8390")

        txt_sub_conta.Items.Add("02500")
        txt_sub_conta.Items.Add("02510")
        txt_sub_conta.Items.Add("03102")
        txt_sub_conta.Items.Add("03112")
        txt_sub_conta.Items.Add("03202")
        txt_sub_conta.Items.Add("03702")
        txt_sub_conta.Items.Add("05402")
        txt_sub_conta.Items.Add("05502")
        txt_sub_conta.Items.Add("10115")
        txt_sub_conta.Items.Add("01115")
    End Sub
    Sub Lista_Req()
        conectar()
        comandoSQL.CommandText = "SELECT * FROM TAB_OS_RELATORIO WHERE OS_SITUACAO = 'ENVIADA'"
        objDataReader = comandoSQL.ExecuteReader
        Lista_Request.Clear()
        Lista_Request.View = View.Details
        Lista_Request.Columns.Add("", 0, HorizontalAlignment.Left)
        Lista_Request.Columns.Add("OS", 60, HorizontalAlignment.Center)
        Lista_Request.Columns.Add("FERRAMENTA", 180, HorizontalAlignment.Center)
        Lista_Request.Columns.Add("POSIÇÃO", 180, HorizontalAlignment.Center)
        Lista_Request.Columns.Add("DATA_ENVIO", 142, HorizontalAlignment.Center)
        Lista_Request.Columns.Add("RESPONSAVEL", 150, HorizontalAlignment.Center)
        Lista_Request.Columns.Add("SITUAÇÃO", 150, HorizontalAlignment.Center)
        While objDataReader.Read()
            Dim ls As New ListViewItem("", 0)
            ls.SubItems.Add(objDataReader("OS_ID").ToString())
            ls.SubItems.Add(objDataReader("OS_FERRAMENTA").ToString())
            ls.SubItems.Add(objDataReader("OS_POSICAO").ToString())
            ls.SubItems.Add(Convert.ToDateTime(objDataReader("OS_DATA_ENVIADA")).ToString("dd/MM/yyyy"))
            ls.SubItems.Add(objDataReader("OS_REG_RESPONSAVEL").ToString())
            ls.SubItems.Add(objDataReader("OS_SITUACAO").ToString())
            Lista_Request.Items.Add(ls)
        End While
        objDataReader.Close()
        desconectar()
    End Sub

    Private Sub txt_caixa_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txt_qtd_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_qtd.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txt_secao_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_secao.KeyPress
        Dim ch As Char = e.KeyChar
        If Not Char.IsDigit(ch) AndAlso ch <> "." AndAlso Not Char.IsControl(ch) Then
            e.Handled = True
        End If
        If ch = "." AndAlso txt_secao.Text.IndexOf(".") <> -1 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txt_op_KeyPress(sender As Object, e As KeyPressEventArgs)
        e.Handled = True
    End Sub

    Private Sub txt_op_qtd_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Sub Resetar()
        txt_OS.Text = ""
        txt_fer.Text = ""
        txt_pos.Text = ""
        txt_desc.Text = ""
        txt_qtd.Text = ""
        txt_tp.Text = ""
        txt_motivo.Items.Clear()
        txt_motivo.Enabled = False
        txt_conta.Text = ""
        txt_sub_conta.Text = ""
        txt_secao.Text = ""
        id_OS = 0
        If id_OS = 0 Then
            conectar()
            comandoSQL.CommandText = " SELECT TOP 1 OS_ID FROM TAB_OS_RELATORIO ORDER BY OS_ID DESC"
            objDataReader = comandoSQL.ExecuteReader
            If objDataReader.Read Then
                COD_FORM = objDataReader(0) + 1
            Else
                COD_FORM = 1
            End If
            objDataReader.Close()
            desconectar()
            txt_OS.Text = COD_FORM
        Else
            COD_FORM = txt_OS.Text
        End If
        Lista_Req()
    End Sub

    Private Sub txt_conta_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_conta.KeyPress
        Dim ch As Char = e.KeyChar
        If Not Char.IsDigit(ch) AndAlso ch <> "." AndAlso Not Char.IsControl(ch) Then
            e.Handled = True
        End If
        If ch = "." AndAlso txt_conta.Text.IndexOf(".") <> -1 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txt_sub_conta_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_sub_conta.KeyPress
        Dim ch As Char = e.KeyChar
        If Not Char.IsDigit(ch) AndAlso ch <> "." AndAlso Not Char.IsControl(ch) Then
            e.Handled = True
        End If
        If ch = "." AndAlso txt_sub_conta.Text.IndexOf(".") <> -1 Then
            e.Handled = True
        End If
    End Sub
    Private Sub txt_limpar_Click(sender As Object, e As EventArgs) Handles txt_limpar.Click
        Resetar()
    End Sub

    Private Sub btn_servico_Click(sender As Object, e As EventArgs) 
        Servico.Show()
    End Sub

    Private Sub btn_relatorio_Click(sender As Object, e As EventArgs) 
        Relatorio.Show()
        Close()
    End Sub

    Private Sub btn_user_MouseHover(sender As Object, e As EventArgs) Handles btn_user.MouseHover
        btn_user.Text = ""
        btn_user.BackgroundImage = Ordem_de_Servico.My.Resources.Resources.Design_sem_nome__7_
    End Sub

    Private Sub btn_user_MouseLeave(sender As Object, e As EventArgs) Handles btn_user.MouseLeave
        btn_user.BackgroundImage = Nothing
        Dim names() As String = UserName.Split(" "c)
        If names.Length >= 2 Then
            Dim initials As String = names(0)(0) & names(names.Length - 1)(0)
            btn_user.Text = initials
        End If
    End Sub

    Private Sub btn_user_Click(sender As Object, e As EventArgs) Handles btn_user.Click
        Close()
        Login.Show()
    End Sub

    Private Sub btn_home_Click(sender As Object, e As EventArgs) Handles btn_home.Click
        Tela_Inicial.Show()
        Close()
    End Sub

    Private Sub txt_tp_SelectedIndexChanged(sender As Object, e As EventArgs) Handles txt_tp.SelectedIndexChanged
        txt_motivo.Enabled = True
        txt_motivo.Items.Clear()
        If txt_tp.SelectedItem.ToString() = "PEÇA NOVA" Then
            txt_motivo.Items.Add("CAVIDADE FECHADA")
            txt_motivo.Items.Add("FTQ")
            txt_motivo.Items.Add("MELHORIA NA FERRAMENTA")
            txt_motivo.Items.Add("PREVENTIVA")
            txt_motivo.Items.Add("CORRETIVA")
            txt_motivo.Items.Add("ESTOQUE")
        Else
            txt_motivo.Items.Add("CAVIDADE FECHADA")
            txt_motivo.Items.Add("FTQ")
            txt_motivo.Items.Add("MELHORIA NA FERRAMENTA")
            txt_motivo.Items.Add("PREVENTIVA")
            txt_motivo.Items.Add("CORRETIVA")
        End If
    End Sub

    Private Sub txt_motivo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_motivo.KeyPress
        e.Handled = True
    End Sub


    Private Sub btn_fechar_Click(sender As Object, e As EventArgs) Handles btn_fechar.Click
        Login.Close()
    End Sub

    Private Sub btn_expandir_Click(sender As Object, e As EventArgs) Handles btn_expandir.Click
        If WindowState = FormWindowState.Maximized Then
            WindowState = FormWindowState.Normal
            btn_expandir.BackgroundImage = Ordem_de_Servico.My.Resources.Resources._7
        Else
            WindowState = FormWindowState.Maximized
            btn_expandir.BackgroundImage = Ordem_de_Servico.My.Resources.Resources._8
        End If
    End Sub

    Private Sub btn_minimizar_Click(sender As Object, e As EventArgs) Handles btn_minimizar.Click
        WindowState = FormWindowState.Minimized
    End Sub
End Class