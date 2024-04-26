Public Class Maquina
    Private Sub Maquina_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Reset()
        Lista()
    End Sub
    Private Sub btn_cadastrar_Click(sender As Object, e As EventArgs) Handles btn_cadastrar.Click
        If txt_maq.Text = "" Then
            MessageBox.Show("Preencha a máquina")
            Return
        End If
        If txt_abre.Text = "" Then
            MessageBox.Show("Preencha a abreviação")
            Return
        End If
        If txt_hr.Text = "" Then
            MessageBox.Show("Preencha a hora semana")
            Return
        End If
        If id_maquina <> 0 Then
            Try
                conectar()
                comandoSQL.CommandText = "UPDATE TAB_OS_MAQUINA SET MAQ_NOME = '" & txt_maq.Text & "', MAQ_ABREVIACAO = '" & txt_abre.Text & "', MAQ_HORA_SEMANA = '" & txt_hr.Text & "' WHERE MAQ_ID = '" & id_maquina & "' "
                comandoSQL.ExecuteNonQuery()
                desconectar()
                MessageBox.Show("Alteração bem-sucedida.")
            Catch ex As Exception
                MessageBox.Show("Ocorreu um erro: " & ex.Message)
            End Try
        Else
            Try
                conectar()
                comandoSQL.CommandText = "INSERT INTO TAB_OS_MAQUINA (MAQ_NOME, MAQ_ABREVIACAO, MAQ_HORA_SEMANA) VALUES ('" & txt_maq.Text & "', '" & txt_abre.Text & "', '" & txt_hr.Text & "')"
                comandoSQL.ExecuteNonQuery()
                desconectar()
                MessageBox.Show("Inserção bem-sucedida.")
            Catch ex As Exception
                MessageBox.Show("Ocorreu um erro: " & ex.Message)
            End Try
        End If
        Reset()
        Lista()
    End Sub

    Sub Reset()
        txt_maq.Text = ""
        txt_abre.Text = ""
        txt_hr.Text = ""
        id_maquina = 0
        btn_cadastrar.Text = "CADASTRAR"
    End Sub

    Sub Lista()
        conectar()
        comandoSQL.CommandText = "SELECT * FROM TAB_OS_MAQUINA"
        objDataReader = comandoSQL.ExecuteReader
        Lista_MAQ.Clear()
        Lista_MAQ.View = View.Details
        Lista_MAQ.Columns.Add("", 0, HorizontalAlignment.Left)
        Lista_MAQ.Columns.Add("ID", 30, HorizontalAlignment.Center)
        Lista_MAQ.Columns.Add("NOME", 120, HorizontalAlignment.Center)
        Lista_MAQ.Columns.Add("ABREVIAÇÃO", 90, HorizontalAlignment.Center)
        Lista_MAQ.Columns.Add("HORA SEMANA", 100, HorizontalAlignment.Center)
        While objDataReader.Read()
            Dim ls As New ListViewItem("", 0)
            ls.SubItems.Add(objDataReader("MAQ_ID").ToString())
            ls.SubItems.Add(objDataReader("MAQ_NOME").ToString())
            ls.SubItems.Add(objDataReader("MAQ_ABREVIACAO").ToString())
            ls.SubItems.Add(objDataReader("MAQ_HORA_SEMANA").ToString().Replace(",", "."))
            Lista_MAQ.Items.Add(ls)
        End While
        objDataReader.Close()
        desconectar()
    End Sub

    Private Sub Lista_MAQ_DoubleClick(sender As Object, e As EventArgs) Handles Lista_MAQ.DoubleClick
        If Lista_MAQ.SelectedItems.Count > 0 Then
            id_maquina = Convert.ToInt32(Lista_MAQ.SelectedItems(0).SubItems(1).Text)
            conectar()
            comandoSQL.CommandText = "SELECT * FROM TAB_OS_MAQUINA WHERE MAQ_ID = '" & id_maquina & "'"
            objDataReader = comandoSQL.ExecuteReader
            If objDataReader.Read() Then
                txt_maq.Text = objDataReader("MAQ_NOME").ToString()
                txt_abre.Text = objDataReader("MAQ_ABREVIACAO").ToString()
                txt_hr.Text = objDataReader("MAQ_HORA_SEMANA").ToString().Replace(",", ".")
                btn_cadastrar.Text = "EDITAR"
            Else
                Reset()
            End If
            objDataReader.Close()
            desconectar()
        End If
    End Sub

    Private Sub txt_hr_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_hr.KeyPress
        Dim ch As Char = e.KeyChar
        If Not Char.IsDigit(ch) AndAlso ch <> "." AndAlso Not Char.IsControl(ch) Then
            e.Handled = True
        End If
        If ch = "." AndAlso txt_hr.Text.IndexOf(".") <> -1 Then
            e.Handled = True
        End If
    End Sub
End Class