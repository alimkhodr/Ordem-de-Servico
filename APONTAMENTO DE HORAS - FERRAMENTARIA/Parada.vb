Public Class Parada
    Private Sub Parada_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cbo_motivo.Items.Add("REFEIÇÃO")
        cbo_motivo.Items.Add("TROCA DE PRIORIDADE")
        cbo_motivo.Items.Add("MANUTENÇÃO CORRETIVA")
        cbo_motivo.Items.Add("MANUTENÇÃO PREVENTIVA")
        cbo_motivo.Items.Add("FALTA DE ENERGIA ELÉTRICA")
        cbo_motivo.Items.Add("FALTA DE PROGRAMA")
        cbo_motivo.Items.Add("FALTA ")
        cbo_motivo.Items.Add("ORGANIZAÇÃO E LIMPEZA")
        cbo_motivo.Items.Add("REUNIÃO OU TREINAMENTO")
    End Sub
    Private Sub btn_parar_Click(sender As Object, e As EventArgs) Handles btn_parar.Click
        If cbo_motivo.Text = "" Then
            MessageBox.Show("Preencha o motivo.")
            Return
        End If
        Try
            conectar()
            objDataReader.Close()
            comandoSQL.CommandText = "INSERT INTO TAB_OP_PARADAS (PARADA_ID_OP, PARADA_DATA_INICIO, PARADA_MOTIVO) VALUES ('" & Servico.id_OP & "', '" & DateTime.Now & "', '" & cbo_motivo.Text & "')"
            comandoSQL.ExecuteReader()
            desconectar()
            MessageBox.Show("Operação parada.")
            Tela_Inicial.Lista()
            Servico.btn_iniciar.Enabled = False
            Servico.btn_parar.Enabled = True
            Servico.btn_parar.Text = "RETOMAR"
            Servico.btn_fim.Enabled = False
            Servico.txt_qtd.Enabled = False
            Close()
        Catch ex As Exception
            MessageBox.Show("Ocorreu um erro: " & ex.Message)
        End Try
    End Sub


End Class
