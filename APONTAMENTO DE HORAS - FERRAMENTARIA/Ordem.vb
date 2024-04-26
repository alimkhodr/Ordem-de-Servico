

Imports System.Globalization
Imports System.IO

Public Class Ordem
    Public COD_FORM As Integer
    Public REGISTRO As String
    Public situacao As String
    Public id_OP As Integer

    Private Sub Ordem_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        permissoes()
        Preencher_cbo()

        If id_OS = 0 Then
            txt_resp.Text = User & " - " & UserName
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
        Lista()

        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("", "")
        DataGridView1.Rows.Clear()
        Dim values As String() = {"MP", "TO", "FR1", "TTA", "TT", "RPL", "RP1", "RCI", "CAM", "CAD", "CUV", "EE", "EEF", "MO"}
        For Each value As String In values
            Dim row As New DataGridViewRow()
            row.CreateCells(DataGridView1)
            row.Cells(0).Value = value
            DataGridView1.Rows.Add(row)
        Next

    End Sub

    Private Sub btn_salvar_Click(sender As Object, e As EventArgs) Handles btn_salvar.Click
        If txt_fer.Text = "" Or txt_externo.Text = "" Or txt_pos.Text = "" Or txt_secao.Text = "" Or txt_resp.Text = "" Or txt_secao.Text = "" Or txt_uni.Text = "" Or txt_tp.Text = "" Or txt_motivo.Text = "" Or txt_qtd.Text = "" Or txt_conta.Text = "" Or txt_sub_conta.Text = "" Or txt_aguardando.Text = "" Then
            MessageBox.Show("Preencha os campos obrigatórios")
            Return
        End If

        Try
            REGISTRO = txt_resp.Text.ToString().Substring(0, txt_resp.Text.ToString().IndexOf(" - "))
        Catch ex As Exception
            MessageBox.Show("Selecione um registro")
            Return
        End Try

        If txt_externo.Text = "SIM" Then
            situacao = "EM REQUISIÇÃO"
        Else
            situacao = "EM ANDAMENTO"
        End If

        If id_OS = 0 Then
            Try
                conectar()
                objDataReader.Close()
                comandoSQL.CommandText = "INSERT INTO TAB_OS_RELATORIO (OS_ID, OS_FERRAMENTA, OS_POSICAO, OS_DESCRICAO, OS_QUANTIDADE, OS_PROJETO, OS_DATA_ENTREGA, OS_UNIDADE, OS_TIPO, OS_MOTIVO, "
                comandoSQL.CommandText &= "OS_CONTA,  OS_SUB_CONTA, OS_SECAO, OS_REG_RESPONSAVEL, OS_SITUACAO, OS_DATA_ENVIADA, OS_DATA_INICIO, OS_AGUARDANDO, OS_CAIXA, OS_QUANTIDADE_FINAL) VALUES ('" & COD_FORM & "', '" & txt_fer.Text & "', '" & txt_pos.Text & "', '" & txt_desc.Text & "', '" & txt_qtd.Text & "', "
                comandoSQL.CommandText &= "'" & txt_projeto.Text & "', '" & txt_dt_ent.Text & "', '" & txt_uni.Text & "', '" & txt_tp.Text & "', '" & txt_motivo.Text & "', '" & txt_conta.Text & "', '" & txt_sub_conta.Text & "', "
                comandoSQL.CommandText &= "'" & txt_secao.Text & "', '" & REGISTRO & "', '" & situacao & "', '" & DateTime.Now & "', '" & DateTime.Now & "', '" & txt_aguardando.Text & "' "
                If String.IsNullOrEmpty(txt_caixa.Text) Then
                    comandoSQL.CommandText &= ", NULL "
                Else
                    comandoSQL.CommandText &= ", '" & txt_caixa.Text & "'"
                End If
                If String.IsNullOrEmpty(txt_qtd_fim.Text) Then
                    comandoSQL.CommandText &= ", NULL)"
                Else
                    comandoSQL.CommandText &= ", '" & txt_qtd_fim.Text & "')"
                End If
                comandoSQL.ExecuteReader()
                desconectar()
                Resetar()
                MessageBox.Show("Ordem criada.")
            Catch ex As Exception
                MessageBox.Show("Ocorreu um erro: " & ex.Message)
            End Try
        Else
            If btn_salvar.Text = "INICIAR" Then
                Try
                    conectar()
                    comandoSQL.CommandText = "UPDATE TAB_OS_RELATORIO SET OS_FERRAMENTA = '" & txt_fer.Text & "', OS_POSICAO = '" & txt_pos.Text & "', OS_DESCRICAO = '" & txt_desc.Text & "', OS_PROJETO = '" & txt_projeto.Text & "', "
                    comandoSQL.CommandText &= "OS_DATA_ENTREGA = '" & txt_dt_ent.Text & "', OS_UNIDADE = '" & txt_uni.Text & "', OS_TIPO = '" & txt_tp.Text & "', OS_MOTIVO = '" & txt_motivo.Text & "', OS_QUANTIDADE = '" & txt_qtd.Text & "', "
                    comandoSQL.CommandText &= "OS_CONTA = '" & txt_conta.Text & "',OS_SUB_CONTA = '" & txt_sub_conta.Text & "',OS_SECAO = '" & txt_secao.Text & "',OS_REG_RESPONSAVEL= '" & REGISTRO & "', OS_SITUACAO = '" & situacao & "', OS_AGUARDANDO = '" & txt_aguardando.Text & "', OS_DATA_INICIO = '" & DateTime.Now & "'  "
                    If String.IsNullOrEmpty(txt_caixa.Text) Then
                        comandoSQL.CommandText &= ", OS_CAIXA = NULL "
                    Else
                        comandoSQL.CommandText &= ", OS_CAIXA = '" & txt_caixa.Text & "'"
                    End If
                    If String.IsNullOrEmpty(txt_qtd_fim.Text) Then
                        comandoSQL.CommandText &= ", OS_QUANTIDADE_FINAL = NULL "
                    Else
                        comandoSQL.CommandText &= ", OS_QUANTIDADE_FINAL = '" & txt_qtd_fim.Text & "'"
                    End If
                    comandoSQL.CommandText &= " WHERE OS_ID = '" & COD_FORM & "'"
                    comandoSQL.ExecuteNonQuery()
                    desconectar()
                    Lista_Request.Visible = False
                    txt_situ.Text = situacao
                    btn_salvar.Text = "SALVAR"
                    MessageBox.Show("Ordem criada.")
                Catch ex As Exception
                    MessageBox.Show("Ocorreu um erro: " & ex.Message)
                End Try
            Else btn_salvar.Text = "SALVAR"
                Try
                    conectar()
                    comandoSQL.CommandText = "UPDATE TAB_OS_RELATORIO SET OS_FERRAMENTA = '" & txt_fer.Text & "', OS_POSICAO = '" & txt_pos.Text & "', OS_DESCRICAO = '" & txt_desc.Text & "', OS_PROJETO = '" & txt_projeto.Text & "', "
                    comandoSQL.CommandText &= "OS_DATA_ENTREGA = '" & txt_dt_ent.Text & "', OS_UNIDADE = '" & txt_uni.Text & "', OS_TIPO = '" & txt_tp.Text & "',  OS_MOTIVO = '" & txt_motivo.Text & "', OS_QUANTIDADE = '" & txt_qtd.Text & "', "
                    comandoSQL.CommandText &= "OS_CONTA = '" & txt_conta.Text & "', OS_SUB_CONTA = '" & txt_sub_conta.Text & "', OS_SECAO = '" & txt_secao.Text & "', OS_REG_RESPONSAVEL= '" & REGISTRO & "', OS_SITUACAO = '" & txt_situ.Text & "', OS_AGUARDANDO = '" & txt_aguardando.Text & "'"
                    If String.IsNullOrEmpty(txt_caixa.Text) Then
                        comandoSQL.CommandText &= ", OS_CAIXA = NULL "
                    Else
                        comandoSQL.CommandText &= ", OS_CAIXA = '" & txt_caixa.Text & "'"
                    End If
                    If String.IsNullOrEmpty(txt_qtd_fim.Text) Then
                        comandoSQL.CommandText &= ", OS_QUANTIDADE_FINAL = NULL "
                    Else
                        comandoSQL.CommandText &= ", OS_QUANTIDADE_FINAL = '" & txt_qtd_fim.Text & "'"
                    End If
                    comandoSQL.CommandText &= " WHERE OS_ID = '" & COD_FORM & "'"
                    comandoSQL.ExecuteNonQuery()
                    desconectar()
                    MessageBox.Show("Alteração bem-sucedida.")
                Catch ex As Exception
                    MessageBox.Show("Ocorreu um erro: " & ex.Message)
                End Try
            End If
        End If
        Lista_Req()
    End Sub

    Sub Preencher_cbo()
        conectar()
        comandoSQL.CommandText = "SELECT FUN_REGISTRO, FUN_NOME FROM FUNCIONARIO WHERE FUN_STATUS = 'ATIVO'"
        objDataReader = comandoSQL.ExecuteReader
        While objDataReader.Read
            txt_resp.Items.Add(objDataReader("FUN_REGISTRO") & " - " & objDataReader("FUN_NOME"))
        End While
        desconectar()

        txt_uni.Items.Clear()
        txt_uni.Items.Add("")
        txt_uni.Items.Add("PC")
        txt_uni.Items.Add("CJ")

        txt_tp.Items.Clear()
        txt_tp.Items.Add("")
        txt_tp.Items.Add("RETRABALHO")
        txt_tp.Items.Add("PEÇA NOVA")

        txt_aguardando.Items.Clear()
        txt_aguardando.Items.Add("")
        txt_aguardando.Items.Add("NÃO")
        txt_aguardando.Items.Add("SIM")

        txt_op.Items.Clear()
        txt_op.Items.Add("")
        txt_op.Items.Add("MP")
        txt_op.Items.Add("TO")
        txt_op.Items.Add("FR1")
        txt_op.Items.Add("TTA")
        txt_op.Items.Add("TT")
        txt_op.Items.Add("RPL")
        txt_op.Items.Add("RP1")
        txt_op.Items.Add("RCI")
        txt_op.Items.Add("CAM")
        txt_op.Items.Add("CAD")
        txt_op.Items.Add("CUV")
        txt_op.Items.Add("EE")
        txt_op.Items.Add("EEF")
        txt_op.Items.Add("MO")

        txt_conta.Items.Clear()
        txt_conta.Items.Add("7000")
        txt_conta.Items.Add("8390")

        txt_sub_conta.Items.Clear()
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

        If txt_tp.Text = "PEÇA NOVA" Then
            txt_motivo.Items.Clear()
            txt_motivo.Items.Add("CAVIDADE FECHADA")
            txt_motivo.Items.Add("FTQ")
            txt_motivo.Items.Add("MELHORIA NA FERRAMENTA")
            txt_motivo.Items.Add("PREVENTIVA")
            txt_motivo.Items.Add("CORRETIVA")
            txt_motivo.Items.Add("ESTOQUE")
        Else
            txt_motivo.Items.Clear()
            txt_motivo.Items.Add("CAVIDADE FECHADA")
            txt_motivo.Items.Add("FTQ")
            txt_motivo.Items.Add("MELHORIA NA FERRAMENTA")
            txt_motivo.Items.Add("PREVENTIVA")
            txt_motivo.Items.Add("CORRETIVA")
        End If

        txt_externo.Items.Clear()
        txt_externo.Items.Add("")
        txt_externo.Items.Add("SIM")
        txt_externo.Items.Add("NÃO")
    End Sub

    Private Sub btn_inserir_Click(sender As Object, e As EventArgs) Handles btn_inserir.Click
        If btn_inserir.Text = "INSERIR" Then
            Dim count As Integer = DataGridView1.SelectedRows.Count
            If count = 0 Then
                MessageBox.Show("Selecione pelo menos uma operação.")
                Return
            End If
            Try
                conectar()
                objDataReader.Close()
                For Each row As DataGridViewRow In DataGridView1.SelectedRows
                    Dim operacao As String = row.Cells(0).Value
                    comandoSQL.CommandText = "INSERT INTO TAB_OS_OPERACAO (OP_NOME_OP, OP_DESCRICAO_OP, OP_ID_OS, OP_DATA_ENTREGA, OP_HORAS_PREV, OP_HORAS_PREV_UNIT, OP_QUANTIDADE) "
                    comandoSQL.CommandText &= "VALUES ('" & operacao & "', '" & txt_op_desc.Text & "', '" & COD_FORM & "', '" & txt_op_dt.Text & "', '" & txt_op_hr.Text & "', '" & txt_op_hr_uni.Text & "', '" & txt_op_qtd.Text & "')"
                    comandoSQL.ExecuteNonQuery()
                Next
                desconectar()
                Lista()
                MessageBox.Show("Inserção bem-sucedida.")
            Catch ex As Exception
                MessageBox.Show("Ocorreu um erro: " & ex.Message)
            End Try
        Else
            If txt_op.Text = "" Then
                MessageBox.Show("Selecione uma operação.")
                Return
            End If
            Try
                conectar()
                objDataReader.Close()
                comandoSQL.CommandText = "UPDATE TAB_OS_OPERACAO SET OP_NOME_OP = '" & txt_op.Text & "', OP_DESCRICAO_OP = '" & txt_op_desc.Text & "', OP_DATA_ENTREGA = '" & txt_op_dt.Text & "', OP_HORAS_PREV = '" & txt_op_hr.Text & "', "
                comandoSQL.CommandText &= "OP_HORAS_PREV_UNIT = '" & txt_op_hr_uni.Text & "', OP_QUANTIDADE = '" & txt_op_qtd.Text & "' WHERE OP_ID = '" & id_OP & "'"
                comandoSQL.ExecuteNonQuery()
                desconectar()
                Lista()
                MessageBox.Show("Alteração bem-sucedida.")
            Catch ex As Exception
                MessageBox.Show("Ocorreu um erro: " & ex.Message)
            End Try
        End If
        resetar_op()
    End Sub

    Sub resetar_op()
        id_OP = 0
        txt_op.Text = ""
        txt_op_desc.Text = ""
        txt_op_dt.Text = ""
        txt_op_hr.Text = "0"
        txt_op_hr_uni.Text = "0"
        txt_op_qtd.Text = "0"

        DataGridView1.Location = New Point(9, 270)
        DataGridView1.Size = New Point(92, 162)
        DataGridView1.Enabled = True
        txt_op.Visible = False

        btn_inserir.Text = "INSERIR"
    End Sub

    Private Sub btn_deletar_Click(sender As Object, e As EventArgs) Handles btn_deletar.Click
        If id_OP = 0 Then
            MessageBox.Show("Selecione uma operação")
            Return
        End If
        Dim result As DialogResult = MessageBox.Show("Tem certeza que deseja excluir esta operação?", "Confirmar exclusão", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            Try
                conectar()
                objDataReader.Close()
                comandoSQL.CommandText = "DELETE TAB_OS_OPERACAO WHERE OP_ID = '" & id_OP & "'"
                comandoSQL.ExecuteReader()
                desconectar()
                id_OP = 0
                Lista()
                MessageBox.Show("Deletado com sucesso.")
            Catch ex As Exception
                MessageBox.Show("Ocorreu um erro: " & ex.Message)
            End Try
            resetar_op()
        Else
            id_OP = 0
        End If
    End Sub

    Sub Lista()
        conectar()
        comandoSQL.CommandText = "SELECT OP_ID, OP_NOME_OP, OP_DESCRICAO_OP, OP_DATA_ENTREGA, OP_HORAS_PREV, OP_HORAS_PREV_UNIT, OP_QUANTIDADE, OP_DATA_FIM, OP_DATA_INICIO FROM TAB_OS_OPERACAO WHERE OP_ID_OS = " & COD_FORM
        objDataReader = comandoSQL.ExecuteReader
        Lista_OP.Clear()
        Lista_OP.View = View.Details
        Lista_OP.Columns.Add("", 0, HorizontalAlignment.Left)
        Lista_OP.Columns.Add("ID", 50, HorizontalAlignment.Center)
        Lista_OP.Columns.Add("OPERAÇÃO", 110, HorizontalAlignment.Center)
        Lista_OP.Columns.Add("DESCRIÇÃO", 110, HorizontalAlignment.Center)
        Lista_OP.Columns.Add("ENTREGA", 105, HorizontalAlignment.Center)
        Lista_OP.Columns.Add("HORAS PREV.", 120, HorizontalAlignment.Center)
        Lista_OP.Columns.Add("HORAS PREV. UNT.", 120, HorizontalAlignment.Center)
        Lista_OP.Columns.Add("QTD", 50, HorizontalAlignment.Center)
        Lista_OP.Columns.Add("SITUAÇÃO", 130, HorizontalAlignment.Center)
        While objDataReader.Read()
            Dim ls As New ListViewItem("", 0)
            ls.SubItems.Add(objDataReader("OP_ID").ToString())
            ls.SubItems.Add(objDataReader("OP_NOME_OP").ToString())
            ls.SubItems.Add(objDataReader("OP_DESCRICAO_OP").ToString())
            ls.SubItems.Add(Convert.ToDateTime(objDataReader("OP_DATA_ENTREGA")).ToString("dd/MM/yyyy"))
            ls.SubItems.Add(objDataReader("OP_HORAS_PREV").ToString().Replace(",", "."))
            ls.SubItems.Add(objDataReader("OP_HORAS_PREV_UNIT").ToString().Replace(",", "."))
            ls.SubItems.Add(objDataReader("OP_QUANTIDADE").ToString())
            If objDataReader("OP_DATA_INICIO") Is DBNull.Value Then
                ls.SubItems.Add("AGUARDANDO")
            Else
                If objDataReader("OP_DATA_FIM") Is DBNull.Value Then
                    ls.SubItems.Add("EM ANDAMENTO")
                Else
                    ls.SubItems.Add("FINALIZADA")
                End If
            End If
            Lista_OP.Items.Add(ls)
        End While
        objDataReader.Close()
        desconectar()
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

    Private Sub Lista_OP_Click(sender As Object, e As EventArgs) Handles Lista_OP.Click
        If Lista_OP.SelectedItems.Count > 0 Then
            id_OP = Convert.ToInt32(Lista_OP.SelectedItems(0).SubItems(1).Text)
        End If
    End Sub


    Private Sub btn_finalizar_Click(sender As Object, e As EventArgs) Handles btn_finalizar.Click
        If txt_qtd_fim.Text = "" Then
            MessageBox.Show("Preencha a quantidade final.")
            Return
        End If
        Dim result As DialogResult = MessageBox.Show("Tem certeza que deseja finalizar esta OS?", "Confirmar finalzação", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            Try
                conectar()
                comandoSQL.CommandText = "UPDATE TAB_OP_PARADAS SET PARADA_DATA_FIM = COALESCE(PARADA_DATA_FIM, GETDATE()) WHERE PARADA_DATA_FIM IS NULL AND PARADA_ID_OP IN (SELECT OP_ID FROM TAB_OS_OPERACAO WHERE OP_ID_OS = " & COD_FORM & ")"
                comandoSQL.ExecuteNonQuery()

                comandoSQL.CommandText = "UPDATE TAB_OS_OPERACAO SET OP_DATA_FIM = COALESCE(OP_DATA_FIM, GETDATE()) WHERE OP_DATA_FIM IS NULL AND OP_ID_OS IN (SELECT OS_ID FROM TAB_OS_RELATORIO WHERE OS_ID = " & COD_FORM & ")"
                comandoSQL.ExecuteNonQuery()

                conectar()
                comandoSQL.CommandText = "UPDATE TAB_OS_RELATORIO SET OS_DATA_FIM = '" & DateTime.Now & "', OS_SITUACAO = 'FINALIZADA', OS_CAIXA = NULL WHERE OS_ID = '" & COD_FORM & "'"
                comandoSQL.ExecuteNonQuery()
                desconectar()
                MessageBox.Show("Finalização bem-sucedida.")
                Resetar()
            Catch ex As Exception
                MessageBox.Show("Ocorreu um erro: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub txt_op_hr_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_op_hr.KeyPress
        Dim ch As Char = e.KeyChar
        If Not Char.IsDigit(ch) AndAlso ch <> "." AndAlso Not Char.IsControl(ch) Then
            e.Handled = True
        End If
        If ch = "." AndAlso txt_op_hr.Text.IndexOf(".") <> -1 Then
            e.Handled = True
        End If

        Dim op_hr As Double
        Dim qtd As Double
        If e.KeyChar = ChrW(Keys.Enter) Then
            If Not Double.TryParse(txt_op_hr.Text, NumberStyles.Any, CultureInfo.InvariantCulture, op_hr) Then
                MessageBox.Show("Por favor, insira um número válido para horas.")
                Return
            End If

            If Not Double.TryParse(txt_op_qtd.Text, NumberStyles.Any, CultureInfo.InvariantCulture, qtd) Then
                MessageBox.Show("Por favor, insira um número válido para quantidade.")
                Return
            End If

            If op_hr <> 0 Then
                Dim resultado As Double = op_hr / qtd
                txt_op_hr_uni.Text = resultado.ToString(CultureInfo.InvariantCulture)
            Else
                MessageBox.Show("O número de horas não pode ser zero.")
            End If
        End If
    End Sub


    Private Sub txt_caixa_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_caixa.KeyPress
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

    Private Sub btn_imprimir_Click(sender As Object, e As EventArgs) Handles btn_imprimir.Click
        Desenho.Show()
    End Sub

    Public Function ObterDados() As DataTable
        Dim dt = New DataTable()
        dt.Columns.Add("barcode", GetType(Byte()))
        dt.Columns.Add("desc")
        dt.Columns.Add("dt_entrega")
        dt.Columns.Add("op")
        dt.Columns.Add("hr_prev")

        For Each item As ListViewItem In Lista_OP.Items
            Dim barcodeImage As Image = GenerateBarcodeImage(item.SubItems(1).Text)
            dt.Rows.Add(ImageToByteArray(barcodeImage), "Descrição: " & item.SubItems(3).Text, "Data Entrega: " & item.SubItems(4).Text, item.SubItems(2).Text, "Horas Prev.: " & item.SubItems(5).Text)
        Next
        Return dt
    End Function

    Public Function GenerateBarcodeImage(ByVal code As String) As Image
        Using bc As New BarcodeLib.Barcode()
            bc.IncludeLabel = True
            bc.LabelFont = New Font("Arial", 35.0F)
            Return bc.Encode(BarcodeLib.TYPE.CODE128, code)
        End Using
    End Function

    Public Function ImageToByteArray(ByVal imageIn As Image) As Byte()
        Dim ms As New MemoryStream()
        imageIn.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
        Return ms.ToArray()
    End Function


    Private Sub txt_op_qtd_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_op_qtd.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If

        Dim op_hr As Double
        Dim qtd As Double
        If e.KeyChar = ChrW(Keys.Enter) Then
            If Not Double.TryParse(txt_op_hr.Text, NumberStyles.Any, CultureInfo.InvariantCulture, op_hr) Then
                MessageBox.Show("Por favor, insira um número válido para horas.")
                Return
            End If

            If Not Double.TryParse(txt_op_qtd.Text, NumberStyles.Any, CultureInfo.InvariantCulture, qtd) Then
                MessageBox.Show("Por favor, insira um número válido para quantidade.")
                Return
            End If

            If op_hr <> 0 Then
                Dim resultado As Double = op_hr / qtd
                txt_op_hr_uni.Text = resultado.ToString(CultureInfo.InvariantCulture)
            Else
                MessageBox.Show("O número de horas não pode ser zero.")
            End If
        End If
    End Sub

    Private Sub Lista_Request_DoubleClick(sender As Object, e As EventArgs) Handles Lista_Request.DoubleClick
        If Lista_Request.SelectedItems.Count > 0 Then
            id_OS = Convert.ToInt32(Lista_Request.SelectedItems(0).SubItems(1).Text)
            conectar()
            comandoSQL.CommandText = "SELECT TAB_OS_RELATORIO.*, FUN_NOME FROM TAB_OS_RELATORIO INNER JOIN FUNCIONARIO ON OS_REG_RESPONSAVEL = FUN_REGISTRO WHERE OS_ID = '" & id_OS & "'"
            objDataReader = comandoSQL.ExecuteReader
            If objDataReader.Read() Then
                txt_OS.Text = objDataReader("OS_ID").ToString()
                txt_fer.Text = objDataReader("OS_FERRAMENTA").ToString()
                txt_pos.Text = objDataReader("OS_POSICAO").ToString()
                txt_desc.Text = objDataReader("OS_DESCRICAO").ToString()
                txt_tp.Text = objDataReader("OS_TIPO").ToString()
                txt_motivo.Text = objDataReader("OS_MOTIVO").ToString()
                txt_dt_ent.Text = objDataReader("OS_DATA_ENTREGA").ToString()
                txt_caixa.Text = objDataReader("OS_CAIXA").ToString()
                txt_uni.Text = objDataReader("OS_UNIDADE").ToString()
                txt_qtd.Text = objDataReader("OS_QUANTIDADE").ToString()
                txt_qtd_fim.Text = objDataReader("OS_QUANTIDADE_FINAL").ToString()
                txt_conta.Text = objDataReader("OS_CONTA").ToString()
                txt_sub_conta.Text = objDataReader("OS_SUB_CONTA").ToString()
                txt_secao.Text = objDataReader("OS_SECAO").ToString()
                txt_resp.Text = objDataReader("OS_REG_RESPONSAVEL").ToString() & " - " & objDataReader("FUN_NOME").ToString()
                txt_projeto.Text = objDataReader("OS_PROJETO").ToString()
                txt_situ.Text = objDataReader("OS_SITUACAO").ToString()
                txt_dt_ini.Text = objDataReader("OS_DATA_INICIO").ToString()
                If objDataReader("OS_SITUACAO").ToString() = "ENVIADA" Then
                    Lista_Request.Visible = True
                    btn_salvar.Text = "INICIAR"
                Else
                    Lista_Request.Visible = False
                End If
            Else
                Resetar()
            End If
            objDataReader.Close()
            desconectar()
        End If
        COD_FORM = txt_OS.Text
        id_OP = 0
    End Sub

    Sub Resetar()
        id_OS = 0
        txt_fer.Text = ""
        txt_externo.Text = ""
        txt_pos.Text = ""
        txt_desc.Text = ""
        txt_tp.Text = ""
        txt_motivo.Items.Clear()
        txt_motivo.Enabled = False
        txt_aguardando.Text = ""
        txt_dt_ent.Text = ""
        txt_caixa.Text = ""
        txt_uni.Text = ""
        txt_qtd.Text = ""
        txt_qtd_fim.Text = ""
        txt_conta.Text = ""
        txt_sub_conta.Text = ""
        txt_secao.Text = ""
        txt_resp.Text = User & " - " & UserName
        txt_projeto.Text = ""
        txt_situ.Text = "EM ANDAMENTO"
        txt_dt_ini.Text = ""
        Lista_Request.Visible = True
        btn_salvar.Text = "INICIAR"
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

    Private Sub txt_qtd_fim_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_qtd_fim.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txt_limpar_Click(sender As Object, e As EventArgs) Handles txt_limpar.Click
        Resetar()
    End Sub

    Private Sub btn_servico_Click(sender As Object, e As EventArgs) Handles btn_servico.Click
        Servico.Show()
    End Sub

    Private Sub btn_relatorio_Click(sender As Object, e As EventArgs) Handles btn_relatorio.Click
        Relatorio.Show()
        Close()
    End Sub

    Private Sub btn_maq_Click(sender As Object, e As EventArgs) Handles btn_maq.Click
        Maquina.Show()
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

    Private Sub Lista_OP_DoubleClick(sender As Object, e As EventArgs) Handles Lista_OP.DoubleClick
        If Lista_OP.SelectedItems.Count > 0 Then
            id_OP = Convert.ToInt32(Lista_OP.SelectedItems(0).SubItems(1).Text)
            txt_op.Text = Lista_OP.SelectedItems(0).SubItems(2).Text
            txt_op_desc.Text = Lista_OP.SelectedItems(0).SubItems(3).Text
            txt_op_dt.Text = Lista_OP.SelectedItems(0).SubItems(4).Text
            txt_op_hr.Text = Lista_OP.SelectedItems(0).SubItems(5).Text
            txt_op_hr_uni.Text = Lista_OP.SelectedItems(0).SubItems(6).Text
            txt_op_qtd.Text = Lista_OP.SelectedItems(0).SubItems(7).Text

            DataGridView1.Location = New Point(9, 305)
            DataGridView1.Size = New Point(92, 126)
            DataGridView1.Enabled = False
            txt_op.Visible = True

            btn_inserir.Text = "EDITAR"
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        resetar_op()
    End Sub

    Private Sub txt_tp_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_tp.KeyPress
        e.Handled = True
    End Sub

    Private Sub txt_aguardando_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_aguardando.KeyPress
        e.Handled = True
    End Sub

    Private Sub txt_uni_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_uni.KeyPress
        e.Handled = True
    End Sub

    Private Sub txt_resp_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_resp.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            txt_resp.Items.Clear()
            Dim nome As String = txt_resp.Text.ToString().Substring(txt_resp.Text.ToString().IndexOf("-") + 1).Trim()
            If txt_resp.Text <> "" Then
                conectar()
                comandoSQL.CommandText = "SELECT FUN_REGISTRO, FUN_NOME FROM FUNCIONARIO WHERE ('" & txt_resp.Text & "' = '' OR FUN_NOME LIKE '%" & nome & "%') AND FUN_STATUS = 'ATIVO'"
                objDataReader = comandoSQL.ExecuteReader
                While objDataReader.Read
                    txt_resp.Items.Add(objDataReader("FUN_REGISTRO") & " - " & objDataReader("FUN_NOME"))
                End While
                desconectar()
            Else
                conectar()
                comandoSQL.CommandText = "SELECT FUN_REGISTRO, FUN_NOME FROM FUNCIONARIO WHERE FUN_STATUS = 'ATIVO'"
                objDataReader = comandoSQL.ExecuteReader
                While objDataReader.Read
                    txt_resp.Items.Add(objDataReader("FUN_REGISTRO") & " - " & objDataReader("FUN_NOME"))
                End While
                desconectar()
            End If
        End If
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

    Private Sub txt_externo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_externo.KeyPress
        e.Handled = True
    End Sub
End Class