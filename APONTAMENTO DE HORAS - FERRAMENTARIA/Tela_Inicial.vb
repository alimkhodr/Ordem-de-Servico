Public Class Tela_Inicial
    Private Sub Tela_Inicial_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        Lista()
        permissoes()
        txt_situacao.Items.Add("")
        txt_situacao.Items.Add("PRODUZINDO")
        txt_situacao.Items.Add("AGUARDANDO")
        txt_situacao.Items.Add("PARADA")
        txt_situacao.Items.Add("OP FINALIZADA")
        txt_situacao.Items.Add("SEM OP")
        txt_situacao.Items.Add("SEM OS")
        txt_situacao.Items.Add("FINALIZADAS")

    End Sub
    Private Sub btn_os_Click(sender As Object, e As EventArgs) Handles btn_os.Click
        id_OS = 0
        If btn_os.Text = "Criar OS" Then
            Ordem.Show()
        Else
            Solicita_Ordem.Show()
        End If
        Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btn_servico.Click
        Servico.Show()
    End Sub

    Private Sub btn_maq_Click(sender As Object, e As EventArgs) Handles btn_maq.Click
        Maquina.Show()
    End Sub

    Sub Lista()
        Try
            conectar()
            comandoSQL.CommandText = "WITH UltimaOperacao AS (SELECT *, ROW_NUMBER() OVER (PARTITION BY OP_ID_OS ORDER BY OP_DATA_INICIO DESC) AS Ordem FROM TAB_OS_OPERACAO) SELECT R.*, UOP.*, PAR.* FROM TAB_OS_RELATORIO R LEFT JOIN UltimaOperacao UOP ON R.OS_ID = UOP.OP_ID_OS AND UOP.Ordem = 1 "
            comandoSQL.CommandText &= "LEFT JOIN (SELECT *, ROW_NUMBER() OVER (PARTITION BY PARADA_ID_OP ORDER BY PARADA_ID DESC) AS RowNum FROM TAB_OP_PARADAS) PAR ON UOP.OP_ID = PAR.PARADA_ID_OP AND PAR.RowNum = 1 "
            comandoSQL.CommandText &= "WHERE ('" & txt_fer.Text & "' = '' OR R.OS_FERRAMENTA LIKE '%" & txt_fer.Text & "%' OR R.OS_POSICAO LIKE '%" & txt_fer.Text & "%' OR R.OS_ID LIKE '%" & txt_fer.Text & "%') "
            comandoSQL.CommandText &= "ORDER BY CASE WHEN R.OS_SITUACAO = 'EM ANDAMENTO' THEN 1 WHEN R.OS_SITUACAO = 'ENVIAVA' THEN 2 ELSE 3 END, R.OS_ID;"
            objDataReader = comandoSQL.ExecuteReader
            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()
            DataGridView1.Columns.Add("OS", "OS")
            DataGridView1.Columns.Add("FERRAMENTA", "FERRAMENTA")
            DataGridView1.Columns.Add("POSIÇÃO", "POSIÇÃO")
            DataGridView1.Columns.Add("INÍCIO", "INÍCIO")
            DataGridView1.Columns.Add("RESPONSÁVEL", "RESPONSÁVEL")
            DataGridView1.Columns.Add("OPERAÇÃO", "OPERAÇÃO")
            DataGridView1.Columns.Add("LOCAL", "LOCAL")
            DataGridView1.Columns.Add("SITUAÇÃO", "SITUAÇÃO")
            While objDataReader.Read()
                Dim situacao As String
                If objDataReader("OS_DATA_INICIO") IsNot DBNull.Value Then
                    If objDataReader("OS_DATA_FIM") Is DBNull.Value Then
                        If objDataReader("OP_NOME_OP") IsNot DBNull.Value Then
                            If objDataReader("OP_MAQUINA") IsNot DBNull.Value Then
                                If objDataReader("OP_DATA_FIM") IsNot DBNull.Value Then
                                    situacao = "OP FINALIZADA"
                                Else
                                    If objDataReader("PARADA_DATA_INICIO") Is DBNull.Value Then
                                        situacao = "PRODUZINDO"
                                    Else
                                        If objDataReader("PARADA_DATA_FIM") Is DBNull.Value Then
                                            situacao = "PARADA"
                                        Else
                                            situacao = "PRODUZINDO"
                                        End If
                                    End If
                                End If
                            Else
                                situacao = "AGUARDANDO"
                            End If
                        Else
                            situacao = "SEM OP"
                        End If
                    Else
                        situacao = "FINALIZADAS"
                    End If
                Else
                    situacao = "SEM OS"
                End If

                Dim row As String() = New String() {
                objDataReader("OS_ID").ToString(),
                objDataReader("OS_FERRAMENTA").ToString(),
                If(objDataReader("OS_POSICAO") Is DBNull.Value, "", objDataReader("OS_POSICAO").ToString()),
                If(objDataReader("OS_DATA_INICIO") Is DBNull.Value, "NÃO INICIADA", Convert.ToDateTime(objDataReader("OS_DATA_INICIO")).ToString("dd/MM/yyyy")),
                 objDataReader("OS_REG_RESPONSAVEL").ToString(),
                 objDataReader("OP_NOME_OP").ToString(),
                objDataReader("OP_MAQUINA").ToString(),
                situacao
            }
                DataGridView1.Rows.Add(row)
            End While
            objDataReader.Close()
            desconectar()
        Catch ex As Exception
            MessageBox.Show("Ocorreu um erro: " & ex.Message)
        End Try
    End Sub

    Private Sub btn_relatorio_Click(sender As Object, e As EventArgs) Handles btn_relatorio.Click
        Relatorio.Show()
        Close()
    End Sub

    Private Sub txt_os_TextChanged(sender As Object, e As EventArgs) Handles txt_fer.TextChanged
        Lista()
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

    Private Sub DataGridView1_DoubleClick(sender As Object, e As EventArgs) Handles DataGridView1.DoubleClick
        If funtipo = 1 Then
            If DataGridView1.SelectedRows.Count > 0 Then
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                id_OS = Convert.ToInt32(selectedRow.Cells(0).Value)
                Try
                    conectar()
                    comandoSQL.CommandText = "SELECT TAB_OS_RELATORIO.*, FUN_NOME FROM TAB_OS_RELATORIO INNER JOIN FUNCIONARIO ON OS_REG_RESPONSAVEL = FUN_REGISTRO WHERE OS_ID = '" & id_OS & "'"
                    objDataReader = comandoSQL.ExecuteReader
                    If objDataReader.Read() Then
                        Ordem.txt_OS.Text = objDataReader("OS_ID").ToString()
                        Ordem.txt_fer.Text = objDataReader("OS_FERRAMENTA").ToString()
                        Ordem.txt_pos.Text = objDataReader("OS_POSICAO").ToString()
                        Ordem.txt_desc.Text = objDataReader("OS_DESCRICAO").ToString()
                        Ordem.txt_tp.Text = objDataReader("OS_TIPO").ToString()
                        Ordem.txt_motivo.Text = objDataReader("OS_MOTIVO").ToString()
                        Ordem.txt_aguardando.Text = objDataReader("OS_AGUARDANDO").ToString()
                        Ordem.txt_dt_ent.Text = objDataReader("OS_DATA_ENTREGA").ToString()
                        Ordem.txt_caixa.Text = objDataReader("OS_CAIXA").ToString()
                        Ordem.txt_uni.Text = objDataReader("OS_UNIDADE").ToString()
                        Ordem.txt_qtd.Text = objDataReader("OS_QUANTIDADE").ToString()
                        Ordem.txt_qtd_fim.Text = objDataReader("OS_QUANTIDADE_FINAL").ToString()
                        Ordem.txt_conta.Text = objDataReader("OS_CONTA").ToString()
                        Ordem.txt_sub_conta.Text = objDataReader("OS_SUB_CONTA").ToString()
                        Ordem.txt_secao.Text = objDataReader("OS_SECAO").ToString()
                        Ordem.txt_resp.Text = objDataReader("OS_REG_RESPONSAVEL").ToString() & " - " & objDataReader("FUN_NOME").ToString()
                        Ordem.txt_projeto.Text = objDataReader("OS_PROJETO").ToString()
                        Ordem.txt_situ.Text = objDataReader("OS_SITUACAO").ToString()
                        Ordem.txt_dt_ini.Text = objDataReader("OS_DATA_INICIO").ToString()
                        Ordem.txt_aguardando.Text = objDataReader("OS_AGUARDANDO").ToString()
                        If objDataReader("OS_SITUACAO").ToString() = "EM ANDAMENTO" Then
                            Ordem.Panel2.Enabled = True
                            Ordem.txt_externo.Enabled = False
                            Ordem.txt_externo.Text = "NÃO"
                            Ordem.Lista_Request.Visible = False
                            Ordem.btn_salvar.Text = "SALVAR"
                        ElseIf objDataReader("OS_SITUACAO").ToString() = "FINALIZADA" Or objDataReader("OS_SITUACAO").ToString() = "EM REQUISIÇÃO" Then
                            Ordem.Panel2.Enabled = False
                        ElseIf objDataReader("OS_SITUACAO").ToString() = "ENVIADA" Then
                            Ordem.txt_externo.Enabled = True
                            Ordem.Panel2.Enabled = True
                            Ordem.Lista_Request.Visible = True
                            Ordem.btn_salvar.Text = "INICIAR"
                        End If
                        Ordem.Lista()
                    End If
                    objDataReader.Close()
                    desconectar()
                Catch ex As Exception
                    MessageBox.Show("Ocorreu um erro: " & ex.Message)
                End Try
            End If
            Ordem.Show()
            Close()
        End If
    End Sub

    Private Sub txt_situacao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles txt_situacao.SelectedIndexChanged
        Dim situacaoSelecionada As String = txt_situacao.SelectedItem.ToString()

        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells("SITUAÇÃO").Value.ToString() = situacaoSelecionada OrElse situacaoSelecionada = "" Then
                row.Visible = True
            Else
                row.Visible = False
            End If
        Next
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Lista()
        txt_situacao.Text = ""
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