Public Class Login

    Private Sub btn_enter_Click(sender As Object, e As EventArgs) Handles btn_enter.Click
        Try
            conectar()
            comandoSQL.CommandText = "SELECT FUN_REGISTRO, FUN_NOME, FUN_SENHA,FUN_CPF FROM FUNCIONARIO WHERE FUN_REGISTRO = '" & txt_registro.Text & "' AND FUN_STATUS = 'ATIVO'"
            objDataReader = comandoSQL.ExecuteReader
            If objDataReader.Read Then
                If txt_senha.Text = objDataReader("FUN_SENHA") OrElse txt_senha.Text = Mid(objDataReader("FUN_CPF"), 1, 6) Then
                    User = objDataReader("FUN_REGISTRO")
                    UserName = objDataReader("FUN_NOME")
                    txt_registro.Text = ""
                    txt_senha.Text = ""
                Else
                    MessageBox.Show("Senha incorreta.")
                    Return
                End If
            Else
                MessageBox.Show("Funcionário não encontrado.")
                Return
            End If
            desconectar()

            conectar()
            comandoSQL.CommandText = "SELECT * FROM TAB_OS_PERMISSAO WHERE PERMISSAO_FUN = " & User
            objDataReader = comandoSQL.ExecuteReader
            If objDataReader.Read() Then
                funtipo = objDataReader("PERMISSAO_TIPO")
            Else
                MessageBox.Show("Funcionário não cadastrado.")
                Return
            End If
            permissoes()
            Tela_Inicial.Show()
            MessageBox.Show("Bem-Vindo(a) " & UserName & " ")
            Hide()
            objDataReader.Close()
            desconectar()
        Catch ex As Exception
            MessageBox.Show("Ocorreu um erro: " & ex.Message)
        End Try
    End Sub

    Private Sub txt_senha_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_senha.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            btn_enter_Click(Nothing, Nothing)
        End If
    End Sub
End Class
