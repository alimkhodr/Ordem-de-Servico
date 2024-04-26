Imports System.Data.SqlClient

Public Class Servico
    Public id_OP As Integer
    Private Sub Servico_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim i As Integer
        i = 0
        While i < 24
            cbo_horai1.Items.Add(i.ToString("00"))
            i = i + 1
        End While

        i = 0
        While i < 60
            cbo_horai2.Items.Add(i.ToString("00"))
            i = i + 1
        End While

        i = 0
        While i < 24
            cbo_horaf1.Items.Add(i.ToString("00"))
            i = i + 1
        End While

        i = 0
        While i < 60
            cbo_horaf2.Items.Add(i.ToString("00"))
            i = i + 1
        End While
        Preencher_maq()
        btn_iniciar.Enabled = False
        btn_parar.Enabled = False
        btn_fim.Enabled = False
        cb_manual.Enabled = False
        cbo_maq.Enabled = False
        txt_qtd.Enabled = False
        Lista_Geral()
        id_OP = 0
    End Sub

    Sub Preencher_maq()
        conectar()
        comandoSQL.CommandText = "SELECT * FROM TAB_OS_MAQUINA"
        objDataReader = comandoSQL.ExecuteReader
        While objDataReader.Read
            cbo_maq.Items.Add(objDataReader("MAQ_ABREVIACAO") & "-" & objDataReader("MAQ_NOME"))
        End While
        objDataReader.Close()
        desconectar()
    End Sub

    Private Sub cb_manual_CheckedChanged(sender As Object, e As EventArgs) Handles cb_manual.CheckedChanged
        If cb_manual.Checked Then
            cbo_horai1.Enabled = True
            cbo_horai2.Enabled = True
            cbo_horaf1.Enabled = True
            cbo_horaf2.Enabled = True
            cbo_dataf.Enabled = True
            cbo_datai.Enabled = True
            btn_iniciar.Enabled = False
            btn_fim.Enabled = True
            txt_qtd.Enabled = True
        Else
            cbo_horai1.Enabled = False
            cbo_horai2.Enabled = False
            cbo_horaf1.Enabled = False
            cbo_horaf2.Enabled = False
            cbo_dataf.Enabled = False
            cbo_datai.Enabled = False
            btn_iniciar.Enabled = True
            btn_fim.Enabled = False
            txt_qtd.Enabled = False
        End If
    End Sub

    Sub Lista_Geral()
        conectar()
        comandoSQL.CommandText = "SELECT OS_FERRAMENTA, OS_POSICAO, OP_ID, OP_DATA_INICIO, OP_MAQUINA FROM TAB_OS_OPERACAO INNER JOIN TAB_OS_RELATORIO ON OP_ID_OS = OS_ID WHERE OP_FUNCIONARIO = '" & User & "' AND OP_DATA_INICIO IS NOT NULL AND OP_DATA_FIM IS NULL AND NOT EXISTS (SELECT 1 FROM TAB_OP_PARADAS AS P WHERE P.PARADA_ID_OP = OP_ID AND P.PARADA_DATA_FIM IS NULL);"
        objDataReader = comandoSQL.ExecuteReader
        Lista_OP.Clear()
        Lista_OP.View = View.Details
        Lista_OP.Columns.Add("", 0, HorizontalAlignment.Left)
        Lista_OP.Columns.Add("ID", 40, HorizontalAlignment.Center)
        Lista_OP.Columns.Add("OP_MAQUINA", 80, HorizontalAlignment.Center)
        Lista_OP.Columns.Add("FERRAMENTA", 80, HorizontalAlignment.Center)
        Lista_OP.Columns.Add("POSIÇÃO", 100, HorizontalAlignment.Center)
        Lista_OP.Columns.Add("INÍCIO", 100, HorizontalAlignment.Center)
        While objDataReader.Read()
            Dim ls As New ListViewItem("", 0)
            ls.SubItems.Add(objDataReader("OP_ID").ToString())
            ls.SubItems.Add(objDataReader("OP_MAQUINA").ToString())
            ls.SubItems.Add(objDataReader("OS_FERRAMENTA").ToString())
            ls.SubItems.Add(objDataReader("OS_POSICAO").ToString())
            ls.SubItems.Add(Convert.ToDateTime(objDataReader("OP_DATA_INICIO")).ToString("dd/MM/yyyy"))
            Lista_OP.Items.Add(ls)
        End While
        objDataReader.Close()
        desconectar()
    End Sub

    Sub Lista()
        conectar()
        comandoSQL.CommandText = "SELECT OS_ID, OS_FERRAMENTA, OS_POSICAO, OP_ID, OP_NOME_OP, OP_DATA_INICIO, OP_MAQUINA FROM TAB_OS_OPERACAO INNER JOIN TAB_OS_RELATORIO ON OP_ID_OS = OS_ID WHERE OP_ID = '" & id_OP & "'"
        objDataReader = comandoSQL.ExecuteReader
        Lista_OP.Clear()
        Lista_OP.View = View.Details
        Lista_OP.Columns.Add("", 0, HorizontalAlignment.Left)
        Lista_OP.Columns.Add("ID", 40, HorizontalAlignment.Center)
        Lista_OP.Columns.Add("OP_MAQUINA", 80, HorizontalAlignment.Center)
        Lista_OP.Columns.Add("FERRAMENTA", 80, HorizontalAlignment.Center)
        Lista_OP.Columns.Add("POSIÇÃO", 100, HorizontalAlignment.Center)
        Lista_OP.Columns.Add("INÍCIO", 100, HorizontalAlignment.Center)
        If objDataReader.Read() Then
            Dim ls As New ListViewItem("", 0)
            ls.SubItems.Add(objDataReader("OP_ID").ToString())
            ls.SubItems.Add(objDataReader("OP_MAQUINA").ToString())
            ls.SubItems.Add(objDataReader("OS_FERRAMENTA").ToString())
            ls.SubItems.Add(objDataReader("OS_POSICAO").ToString())
            If (objDataReader("OP_DATA_INICIO") Is DBNull.Value) Then
                ls.SubItems.Add("NÃO INICIADA")
            Else
                ls.SubItems.Add(Convert.ToDateTime(objDataReader("OP_DATA_INICIO")).ToString("dd/MM/yyyy"))
            End If
            Lista_OP.Items.Add(ls)
            cbo_maq.Text = objDataReader("OP_MAQUINA").ToString()
        End If
        objDataReader.Close()
        desconectar()
    End Sub

    Private Sub btn_iniciar_Click(sender As Object, e As EventArgs) Handles btn_iniciar.Click
        If cbo_maq.Text = "" Then
            MessageBox.Show("Selecione a máquina.")
            Return
        End If
        If txt_caixa.Text = "0" Or txt_caixa.Text = "" Then
            MessageBox.Show("Selecione a caixa.")
            Return
        End If
        Try
            conectar()
            comandoSQL.CommandText = "UPDATE TAB_OS_OPERACAO SET OP_FUNCIONARIO = " & User & ", OP_MAQUINA = '" & cbo_maq.Text & "', OP_DATA_INICIO = GETDATE WHERE OP_ID = " & id_OP
            comandoSQL.ExecuteNonQuery()
            desconectar()
            btn_iniciar.Enabled = False
            btn_parar.Enabled = True
            txt_qtd.Enabled = True
            btn_fim.Enabled = True
            cb_manual.Enabled = False
            cbo_maq.Enabled = True
            MessageBox.Show("Operação iniciada.")
        Catch ex As Exception
            MessageBox.Show("Ocorreu um erro: " & ex.Message)
        End Try
        Try
            conectar()
            comandoSQL.CommandText = "UPDATE TAB_OS_RELATORIO SET OS_CAIXA = " & txt_caixa.Text & " WHERE OS_ID = " & id_OS
            comandoSQL.ExecuteNonQuery()
            desconectar()
        Catch ex As Exception
            MessageBox.Show("Ocorreu um erro: " & ex.Message)
        End Try
        Tela_Inicial.Lista()
    End Sub

    Private Sub btn_fim_Click(sender As Object, e As EventArgs) Handles btn_fim.Click
        If cbo_maq.Text = "" Then
            MessageBox.Show("Máquina não preenchida.")
            Return
        End If
        If txt_qtd.Text = "" Then
            MessageBox.Show("Quantidade não preenchida.")
            Return
        End If
        Dim result As DialogResult = MessageBox.Show("Tem certeza que deseja finalizar esta operação?", "Confirmar finalzação", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            If cb_manual.Checked Then
                If cbo_datai.Text = "" Or cbo_dataf.Text = "" Or cbo_horai1.Text = "" Or cbo_horaf1.Text = "" Or cbo_horai2.Text = "" Or cbo_horaf2.Text = "" Then
                    MessageBox.Show("Data não preenchida.")
                    Return
                End If
                If txt_qtd.Text = "0" Then
                    MessageBox.Show("Caixa não preenchida.")
                    Return
                End If

                Dim datai As Date = Date.Parse(cbo_datai.Text & " " & cbo_horai1.Text & " : " & cbo_horai2.Text & " : 00")
                Dim dataf As Date = Date.Parse(cbo_dataf.Text & " " & cbo_horaf1.Text & " : " & cbo_horaf2.Text & " : 00")
                Try
                    conectar()
                    comandoSQL.CommandText = "UPDATE TAB_OS_OPERACAO SET OP_FUNCIONARIO = " & User & ", OP_MAQUINA = '" & cbo_maq.Text & "', OP_QUANTIDADE_FINAL = '" & txt_qtd.Text & "', OP_DATA_INICIO = '" & datai & "', OP_DATA_FIM = '" & dataf & "' WHERE OP_ID = " & id_OP
                    comandoSQL.ExecuteNonQuery()
                    desconectar()
                    MessageBox.Show("Operação finalizada.")
                    btn_parar.Enabled = False
                    btn_fim.Enabled = False
                    txt_qtd.Enabled = False
                Catch ex As Exception
                    MessageBox.Show("Ocorreu um erro: " & ex.Message)
                End Try
                Try
                    conectar()
                    comandoSQL.CommandText = "UPDATE TAB_OS_RELATORIO SET OS_CAIXA = " & txt_caixa.Text & " WHERE OS_ID = " & id_OS
                    comandoSQL.ExecuteNonQuery()
                    desconectar()
                Catch ex As Exception
                    MessageBox.Show("Ocorreu um erro: " & ex.Message)
                End Try
            Else
                Try
                    conectar()
                    comandoSQL.CommandText = "UPDATE TAB_OS_OPERACAO SET OP_FUNCIONARIO = " & User & ", OP_MAQUINA = '" & cbo_maq.Text & "', OP_QUANTIDADE_FINAL = " & txt_qtd.Text & ", OP_DATA_FIM = GETDATE WHERE OP_ID = " & id_OP
                    comandoSQL.ExecuteNonQuery()
                    desconectar()
                    MessageBox.Show("Operação finalizada.")
                    btn_parar.Enabled = False
                    btn_fim.Enabled = False
                    txt_qtd.Enabled = False
                Catch ex As Exception
                    MessageBox.Show("Ocorreu um erro: " & ex.Message)
                End Try
            End If
            Close()
        End If
        Tela_Inicial.Lista()
    End Sub

    Private Sub btn_parar_Click(sender As Object, e As EventArgs) Handles btn_parar.Click
        If btn_parar.Text = "PARAR" Then
            Parada.Show()
        Else
            Dim ultima_parada As Integer
            Try
                conectar()
                comandoSQL.CommandText = "SELECT TOP 1 * FROM TAB_OP_PARADAS WHERE PARADA_ID_OP = '" & id_OP & "' ORDER BY PARADA_DATA_INICIO DESC;"
                objDataReader = comandoSQL.ExecuteReader
                If objDataReader.Read Then
                    ultima_parada = objDataReader("PARADA_ID")
                End If
                objDataReader.Close()

                comandoSQL.CommandText = "UPDATE TAB_OP_PARADAS SET PARADA_DATA_FIM = '" & DateTime.Now & "' WHERE PARADA_ID = '" & ultima_parada & "'"
                comandoSQL.ExecuteNonQuery()

                comandoSQL.CommandText = "UPDATE TAB_OS_OPERACAO SET OP_FUNCIONARIO = '" & User & "' WHERE OP_ID = '" & id_OP & "'"
                comandoSQL.ExecuteNonQuery()

                desconectar()
                btn_iniciar.Enabled = False
                btn_parar.Enabled = True
                btn_parar.Text = "PARAR"
                btn_fim.Enabled = True
                txt_qtd.Enabled = True
                MessageBox.Show("Operação retomada.")
                Tela_Inicial.Lista()
            Catch ex As Exception
                MessageBox.Show("Ocorreu um erro: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub txt_op_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_op.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
        If e.KeyChar = ChrW(Keys.Enter) Then
            Try
                id_OP = txt_op.Text
                conectar()
                comandoSQL.CommandText = "SELECT TOP 1 OP_ID, OP_ID_OS, OP_NOME_OP, OP_MAQUINA, OP_QUANTIDADE_FINAL, OS_DATA_FIM, OP_DATA_INICIO, OP_DATA_FIM, OP_DATA_FIM,PARADA_ID, PARADA_DATA_INICIO, PARADA_DATA_FIM, OS_CAIXA FROM TAB_OS_OPERACAO INNER JOIN TAB_OS_RELATORIO ON OP_ID_OS = OS_ID LEFT JOIN TAB_OP_PARADAS ON OP_ID = PARADA_ID_OP WHERE OP_ID = '" & id_OP & "' ORDER BY PARADA_ID DESC"
                objDataReader = comandoSQL.ExecuteReader
                If objDataReader.Read Then
                    id_OS = objDataReader("OP_ID_OS")
                    txt_qtd.Text = objDataReader("OP_QUANTIDADE_FINAL").ToString
                    txt_caixa.Text = objDataReader("OS_CAIXA").ToString

                    If Not objDataReader("OP_MAQUINA") Is DBNull.Value Then
                        cbo_maq.Text = objDataReader("OP_MAQUINA").ToString()
                    Else
                        cbo_maq.Text = ""
                    End If

                    If Not objDataReader("OP_DATA_INICIO") Is DBNull.Value Then
                        cbo_datai.Text = CType(objDataReader("OP_DATA_INICIO"), DateTime).ToString("dd/MM/yyyy")
                        cbo_horai1.Text = CType(objDataReader("OP_DATA_INICIO"), DateTime).ToString("HH")
                        cbo_horai2.Text = CType(objDataReader("OP_DATA_INICIO"), DateTime).ToString("mm")
                    Else
                        cbo_datai.Text = ""
                        cbo_horai1.Text = ""
                        cbo_horai2.Text = ""
                    End If

                    If Not objDataReader("OP_DATA_FIM") Is DBNull.Value Then
                        cbo_dataf.Text = CType(objDataReader("OP_DATA_FIM"), DateTime).ToString("dd/MM/yyyy")
                        cbo_horaf1.Text = CType(objDataReader("OP_DATA_FIM"), DateTime).ToString("HH")
                        cbo_horaf2.Text = CType(objDataReader("OP_DATA_FIM"), DateTime).ToString("mm")
                    Else
                        cbo_dataf.Text = ""
                        cbo_horaf1.Text = ""
                        cbo_horaf2.Text = ""
                    End If

                    If objDataReader("OS_DATA_FIM") IsNot DBNull.Value OrElse objDataReader("OS_DATA_FIM") Is Nothing Then
                        Lista()
                        btn_iniciar.Enabled = False
                        btn_parar.Enabled = False
                        btn_fim.Enabled = False
                        txt_qtd.Enabled = False
                        cb_manual.Enabled = False
                        cbo_maq.Enabled = False
                        MessageBox.Show("OS já finalizada.") 'ordem finalizada
                        Return
                    Else
                        If Not objDataReader.IsDBNull(objDataReader.GetOrdinal("OP_DATA_INICIO")) And Not objDataReader.IsDBNull(objDataReader.GetOrdinal("OP_DATA_FIM")) Then 'operação finalizada
                            MessageBox.Show("Operação finalizada")
                            btn_iniciar.Enabled = False
                            btn_parar.Enabled = False
                            btn_fim.Enabled = False
                            txt_qtd.Enabled = False
                            cb_manual.Enabled = False
                            cbo_maq.Enabled = False
                        ElseIf Not objDataReader.IsDBNull(objDataReader.GetOrdinal("PARADA_DATA_INICIO")) And objDataReader.IsDBNull(objDataReader.GetOrdinal("PARADA_DATA_FIM")) Then 'operação parada
                            btn_iniciar.Enabled = False
                            btn_parar.Enabled = True
                            btn_parar.Text = "RETOMAR"
                            btn_fim.Enabled = False
                            txt_qtd.Enabled = False
                            cb_manual.Enabled = False
                            cbo_maq.Enabled = False
                        ElseIf Not objDataReader.IsDBNull(objDataReader.GetOrdinal("OP_DATA_INICIO")) And objDataReader.IsDBNull(objDataReader.GetOrdinal("OP_DATA_FIM")) Then 'operação em andamendo
                            btn_iniciar.Enabled = False
                            btn_parar.Enabled = True
                            btn_fim.Enabled = True
                            txt_qtd.Enabled = True
                            cb_manual.Enabled = False
                            cbo_maq.Enabled = True
                        ElseIf objDataReader.IsDBNull(objDataReader.GetOrdinal("OP_DATA_INICIO")) And objDataReader.IsDBNull(objDataReader.GetOrdinal("OP_DATA_FIM")) Then 'operação não iniciada
                            btn_iniciar.Enabled = True
                            btn_parar.Enabled = False
                            btn_fim.Enabled = False
                            txt_qtd.Enabled = False
                            cb_manual.Enabled = True
                            cbo_maq.Enabled = True
                        End If
                    End If
                    Lista()
                Else
                    btn_iniciar.Enabled = False
                    btn_parar.Enabled = False
                    btn_fim.Enabled = False
                    txt_qtd.Enabled = False
                    Lista_OP.Clear()
                    MessageBox.Show("Operação não encontrada")
                End If
                objDataReader.Close()
                desconectar()
            Catch ex As Exception
                MessageBox.Show("Ocorreu um erro: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub txt_caixa_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_caixa.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub Lista_OP_DoubleClick(sender As Object, e As EventArgs) Handles Lista_OP.DoubleClick
        If Lista_OP.SelectedItems.Count > 0 Then
            Try
                id_OP = Convert.ToInt32(Lista_OP.SelectedItems(0).SubItems(1).Text)
                conectar()
                comandoSQL.CommandText = "SELECT TOP 1 OP_ID, OP_ID_OS, OP_NOME_OP, OP_MAQUINA, OP_QUANTIDADE_FINAL, OS_DATA_FIM, OP_DATA_INICIO, OP_DATA_FIM, OP_DATA_FIM,PARADA_ID, PARADA_DATA_INICIO, PARADA_DATA_FIM, OS_CAIXA FROM TAB_OS_OPERACAO INNER JOIN TAB_OS_RELATORIO ON OP_ID_OS = OS_ID LEFT JOIN TAB_OP_PARADAS ON OP_ID = PARADA_ID_OP WHERE OP_ID = '" & id_OP & "' ORDER BY PARADA_ID DESC"
                objDataReader = comandoSQL.ExecuteReader
                If objDataReader.Read Then
                    id_OS = objDataReader("OP_ID_OS")
                    txt_qtd.Text = objDataReader("OP_QUANTIDADE_FINAL").ToString
                    txt_caixa.Text = objDataReader("OS_CAIXA").ToString
                    If objDataReader("OS_DATA_FIM") IsNot DBNull.Value OrElse objDataReader("OS_DATA_FIM") Is Nothing Then
                        Lista()
                        btn_iniciar.Enabled = False
                        btn_parar.Enabled = False
                        btn_fim.Enabled = False
                        txt_qtd.Enabled = False
                        cb_manual.Enabled = False
                        cbo_maq.Enabled = False
                        MessageBox.Show("OS já finalizada.")
                        Return
                    Else
                        If Not objDataReader.IsDBNull(objDataReader.GetOrdinal("OP_DATA_INICIO")) And Not objDataReader.IsDBNull(objDataReader.GetOrdinal("OP_DATA_FIM")) Then
                            MessageBox.Show("Operação finalizada")
                            btn_iniciar.Enabled = True
                            btn_parar.Enabled = False
                            btn_fim.Enabled = False
                            txt_qtd.Enabled = False
                            cb_manual.Enabled = False
                            cbo_maq.Enabled = False
                        ElseIf Not objDataReader.IsDBNull(objDataReader.GetOrdinal("PARADA_DATA_INICIO")) And objDataReader.IsDBNull(objDataReader.GetOrdinal("PARADA_DATA_FIM")) Then
                            btn_iniciar.Enabled = False
                            btn_parar.Enabled = True
                            btn_fim.Enabled = False
                            txt_qtd.Enabled = False
                            cb_manual.Enabled = False
                            cbo_maq.Enabled = True
                        ElseIf Not objDataReader.IsDBNull(objDataReader.GetOrdinal("OP_DATA_INICIO")) And objDataReader.IsDBNull(objDataReader.GetOrdinal("OP_DATA_FIM")) Then
                            btn_iniciar.Enabled = False
                            btn_parar.Enabled = True
                            btn_fim.Enabled = True
                            txt_qtd.Enabled = True
                            cb_manual.Enabled = False
                            cbo_maq.Enabled = True
                        ElseIf objDataReader.IsDBNull(objDataReader.GetOrdinal("OP_DATA_INICIO")) And objDataReader.IsDBNull(objDataReader.GetOrdinal("OP_DATA_FIM")) Then
                            btn_iniciar.Enabled = True
                            btn_parar.Enabled = False
                            btn_fim.Enabled = False
                            txt_qtd.Enabled = False
                            cb_manual.Enabled = True
                            cbo_maq.Enabled = True
                        End If
                    End If

                    If Not objDataReader("OP_MAQUINA") Is DBNull.Value Then
                        cbo_maq.Text = objDataReader("OP_MAQUINA").ToString()
                    End If

                    If Not objDataReader("OP_DATA_INICIO") Is DBNull.Value Then
                        cbo_datai.Text = CType(objDataReader("OP_DATA_INICIO"), DateTime).ToString("dd/MM/yyyy")
                        cbo_horai1.Text = CType(objDataReader("OP_DATA_INICIO"), DateTime).ToString("HH")
                        cbo_horai2.Text = CType(objDataReader("OP_DATA_INICIO"), DateTime).ToString("mm")
                    End If

                    If Not objDataReader("OP_DATA_FIM") Is DBNull.Value Then
                        cbo_dataf.Text = CType(objDataReader("OP_DATA_FIM"), DateTime).ToString("dd/MM/yyyy")
                        cbo_horaf1.Text = CType(objDataReader("OP_DATA_FIM"), DateTime).ToString("HH")
                        cbo_horaf2.Text = CType(objDataReader("OP_DATA_FIM"), DateTime).ToString("mm")
                    End If
                    Lista()
                End If
                objDataReader.Close()
                desconectar()
            Catch ex As Exception
                MessageBox.Show("Ocorreu um erro: " & ex.Message)
            End Try
        End If
    End Sub
End Class