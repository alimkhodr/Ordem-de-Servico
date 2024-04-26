Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Microsoft.ReportingServices.Rendering.ExcelRenderer

Public Class Relatorio

    Private Sub Relatorio_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Consulta()
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        txt_op.Items.Clear()
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
        Preencher_maq()
        permissoes()

        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("OP", "OP")
        DataGridView1.Columns.Add("HORAS PREV.", "HORAS PREV.")
        DataGridView1.Columns.Add("HORAS REAIS", "HORAS REAIS")
    End Sub
    Sub Preencher_maq()
        conectar()
        comandoSQL.CommandText = "SELECT * FROM TAB_OS_MAQUINA"
        objDataReader = comandoSQL.ExecuteReader
        While objDataReader.Read
            cbo_maq.Items.Add(objDataReader("MAQ_ABREVIACAO") & "-" & objDataReader("MAQ_NOME"))
        End While
        desconectar()
    End Sub
    Sub Consulta()
        Dim deferenca_horas As Decimal = 0
        Dim horastotais As Decimal = 0
        Try
            conectar()
            comandoSQL.CommandText = "SELECT OS_ID, OS_FERRAMENTA, OS_POSICAO, OS_SECAO, OS_PROJETO, OS_CONTA, OS_SUB_CONTA, OS_REG_RESPONSAVEL, OS_TIPO, OS_DATA_INICIO, OS_QUANTIDADE, "
            comandoSQL.CommandText &= "OS_QUANTIDADE_FINAL, OP_ID, OP_NOME_OP, OP_FUNCIONARIO, OP_DATA_INICIO, OP_DATA_FIM, OP_MAQUINA, OP_HORAS_PREV, OP_QUANTIDADE, OP_QUANTIDADE_FINAL,"
            comandoSQL.CommandText &= "DATEDIFF(MINUTE, OP_DATA_INICIO, COALESCE(OP_DATA_FIM, GETDATE())) / 60.0 As HORAS_GERADAS, "
            comandoSQL.CommandText &= "(SELECT COALESCE(SUM(DATEDIFF(MINUTE, PARADA_DATA_INICIO, COALESCE(PARADA_DATA_FIM, GETDATE()))), 0) / 60.0 FROM TAB_OP_PARADAS WHERE PARADA_ID_OP = OP_ID) As SOMA_PARADA "
            comandoSQL.CommandText &= "FROM TAB_OS_RELATORIO INNER JOIN TAB_OS_OPERACAO On OS_ID = OP_ID_OS "
            comandoSQL.CommandText &= "WHERE OP_DATA_INICIO Is Not NULL And ('" & txt_fer.Text & "' = '' OR OS_FERRAMENTA LIKE '%" & txt_fer.Text & "%' OR OS_POSICAO LIKE '%" & txt_fer.Text & "%') AND ('" & txt_os.Text & "' = '' OR OS_ID = '" & txt_os.Text & "') "
            comandoSQL.CommandText &= "AND ('" & txt_data.Text & "' = '  /  /' OR CONVERT(DATE, OP_DATA_INICIO, 103) >= CONVERT(DATE, '" & txt_data.Text & "', 103)) AND ('" & txt_data_fim.Text & "' = '  /  /' OR CONVERT(DATE, OP_DATA_FIM, 103) <= CONVERT(DATE, '" & txt_data_fim.Text & "', 103)) "
            comandoSQL.CommandText &= "AND ('" & txt_op.Text & "' = '' OR OP_NOME_OP = '" & txt_op.Text & "')  AND ('" & cbo_maq.Text & "' = '' OR OP_MAQUINA = '" & cbo_maq.Text & "') "
            comandoSQL.CommandText &= "GROUP BY OS_ID, OS_FERRAMENTA, OS_SECAO, OS_PROJETO, OS_CONTA, OS_SUB_CONTA, OS_REG_RESPONSAVEL, OS_TIPO, OS_DATA_INICIO, OS_QUANTIDADE, "
            comandoSQL.CommandText &= "OS_QUANTIDADE_FINAL, OP_ID, OP_NOME_OP, OP_FUNCIONARIO, OP_DATA_INICIO, OP_DATA_FIM, OP_MAQUINA, "
            comandoSQL.CommandText &= "OP_QUANTIDADE, OP_QUANTIDADE_FINAL, OP_HORAS_PREV, OS_POSICAO ORDER BY OP_ID"
            objDataReader = comandoSQL.ExecuteReader

            Lista_OS.Clear()
            Lista_OS.View = View.Details
            Lista_OS.Columns.Add("ORDEM", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("FERRAMENTA", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("POSIÇÃO", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("PROJETO", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("SEÇÃO", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("CONTA", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("SUB_CONTA", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("REG_RESPONSAVEL", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("TIPO", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("DATA INICIO", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("QUANTIDADE", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("QUANTIDADE FINAL", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("OPERAÇÃO", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("NOME", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("FUNCIONÁRIO", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("DATA INICIO", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("DATA FIM", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("MAQUINA", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("QUANTIDADE", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("QUANTIDADE FINAL", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("HORAS PROGRAMADAS", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("HORAS REAIS", 60, HorizontalAlignment.Center)

            While objDataReader.Read()
                Dim horas_geradas As Decimal = If(objDataReader("HORAS_GERADAS") IsNot DBNull.Value, Convert.ToDecimal(objDataReader("HORAS_GERADAS")), 0)
                Dim horas_parada As Decimal = If(objDataReader("SOMA_PARADA") IsNot DBNull.Value, Convert.ToDecimal(objDataReader("SOMA_PARADA")), 0)
                deferenca_horas = Math.Round(horas_geradas - horas_parada, 2)
                Dim ls As New ListViewItem(objDataReader("OS_ID").ToString())
                ls.SubItems.Add(objDataReader("OS_FERRAMENTA").ToString())
                ls.SubItems.Add(objDataReader("OS_POSICAO").ToString())
                ls.SubItems.Add(objDataReader("OS_PROJETO").ToString())
                ls.SubItems.Add(objDataReader("OS_SECAO").ToString())
                ls.SubItems.Add(objDataReader("OS_CONTA").ToString())
                ls.SubItems.Add(objDataReader("OS_SUB_CONTA").ToString())
                ls.SubItems.Add(objDataReader("OS_REG_RESPONSAVEL").ToString())
                ls.SubItems.Add(objDataReader("OS_TIPO").ToString())
                ls.SubItems.Add(objDataReader("OS_DATA_INICIO").ToString())
                ls.SubItems.Add(objDataReader("OS_QUANTIDADE").ToString())
                ls.SubItems.Add(objDataReader("OS_QUANTIDADE_FINAL").ToString())
                ls.SubItems.Add(objDataReader("OP_ID").ToString())
                ls.SubItems.Add(objDataReader("OP_NOME_OP").ToString())
                ls.SubItems.Add(objDataReader("OP_FUNCIONARIO").ToString())
                ls.SubItems.Add(objDataReader("OP_DATA_INICIO").ToString())
                ls.SubItems.Add(objDataReader("OP_DATA_FIM").ToString())
                ls.SubItems.Add(objDataReader("OP_MAQUINA").ToString())
                ls.SubItems.Add(objDataReader("OP_QUANTIDADE").ToString())
                ls.SubItems.Add(objDataReader("OP_QUANTIDADE_FINAL").ToString())
                ls.SubItems.Add(objDataReader("OP_HORAS_PREV").ToString())
                ls.SubItems.Add(deferenca_horas)
                Lista_OS.Items.Add(ls)
            End While
            objDataReader.Close()

            comandoSQL.CommandText = "SELECT OS_ID, OS_FERRAMENTA, OS_POSICAO, OP_ID, OP_NOME_OP, OP_DATA_INICIO, OP_DATA_FIM, PARADA_MOTIVO,PARADA_DATA_INICIO,PARADA_DATA_FIM, "
            comandoSQL.CommandText &= "Case WHEN PARADA_MOTIVO = 'FIM DE TURNO' OR PARADA_MOTIVO = 'REFEIÇÃO' OR PARADA_MOTIVO = 'TROCA DE PRIORIDADE' OR PARADA_MOTIVO = 'FALTA DE ENERGIA' "
            comandoSQL.CommandText &= "THEN 0 ELSE COALESCE(SUM(DATEDIFF(MINUTE, PARADA_DATA_INICIO, COALESCE(PARADA_DATA_FIM, GETDATE()))), 0) / 60.0 END AS HORAS_PARADA "
            comandoSQL.CommandText &= "From TAB_OS_RELATORIO INNER JOIN TAB_OS_OPERACAO On OS_ID = OP_ID_OS LEFT JOIN TAB_OP_PARADAS On OP_ID = PARADA_ID_OP "
            comandoSQL.CommandText &= "WHERE OP_DATA_INICIO Is Not NULL AND EXISTS (SELECT 1 FROM TAB_OP_PARADAS WHERE PARADA_ID_OP = OP_ID AND PARADA_DATA_FIM IS NULL) And ('" & txt_fer.Text & "' = '' OR OS_FERRAMENTA LIKE '%" & txt_fer.Text & "%' OR OS_POSICAO LIKE '%" & txt_fer.Text & "%') AND ('" & txt_os.Text & "' = '' OR OS_ID = '" & txt_os.Text & "') "
            comandoSQL.CommandText &= "AND ('" & txt_data.Text & "' = '  /  /' OR CONVERT(DATE, OP_DATA_INICIO, 103) >= CONVERT(DATE, '" & txt_data.Text & "', 103)) AND ('" & txt_data_fim.Text & "' = '  /  /' OR CONVERT(DATE, OP_DATA_FIM, 103) <= CONVERT(DATE, '" & txt_data_fim.Text & "', 103)) "
            comandoSQL.CommandText &= "AND ('" & txt_op.Text & "' = '' OR OP_NOME_OP = '" & txt_op.Text & "')  AND ('" & cbo_maq.Text & "' = '' OR OP_MAQUINA = '" & cbo_maq.Text & "') "
            comandoSQL.CommandText &= "GROUP BY OS_ID, OS_FERRAMENTA, OP_ID, OP_NOME_OP, OP_DATA_INICIO, OP_DATA_FIM, OS_POSICAO, PARADA_MOTIVO,PARADA_DATA_INICIO,PARADA_DATA_FIM ORDER BY OP_ID; "
            objDataReader = comandoSQL.ExecuteReader

            Lista_Parada.Clear()
            Lista_Parada.View = View.Details
            Lista_Parada.Columns.Add("ORDEM", 60, HorizontalAlignment.Center)
            Lista_Parada.Columns.Add("FERRAMENTA", 60, HorizontalAlignment.Center)
            Lista_Parada.Columns.Add("POSIÇÃO", 60, HorizontalAlignment.Center)
            Lista_Parada.Columns.Add("OPERAÇÃO", 60, HorizontalAlignment.Center)
            Lista_Parada.Columns.Add("NOME", 60, HorizontalAlignment.Center)
            Lista_Parada.Columns.Add("DATA INICIO", 60, HorizontalAlignment.Center)
            Lista_Parada.Columns.Add("DATA FIM", 60, HorizontalAlignment.Center)
            Lista_Parada.Columns.Add("MOTIVO PARADA", 60, HorizontalAlignment.Center)
            Lista_Parada.Columns.Add("INICIO PARADA", 60, HorizontalAlignment.Center)
            Lista_Parada.Columns.Add("FIM PARADA", 60, HorizontalAlignment.Center)
            Lista_Parada.Columns.Add("HORAS PARADA", 60, HorizontalAlignment.Center)

            While objDataReader.Read()
                Dim ls As New ListViewItem(objDataReader("OS_ID").ToString())
                ls.SubItems.Add(objDataReader("OS_FERRAMENTA").ToString())
                ls.SubItems.Add(objDataReader("OS_POSICAO").ToString())
                ls.SubItems.Add(objDataReader("OP_ID").ToString())
                ls.SubItems.Add(objDataReader("OP_NOME_OP").ToString())
                ls.SubItems.Add(objDataReader("OP_DATA_INICIO").ToString())
                ls.SubItems.Add(objDataReader("OP_DATA_FIM").ToString())
                ls.SubItems.Add(objDataReader("PARADA_MOTIVO").ToString())
                ls.SubItems.Add(objDataReader("PARADA_DATA_INICIO").ToString())
                ls.SubItems.Add(objDataReader("PARADA_DATA_FIM").ToString())
                ls.SubItems.Add(objDataReader("HORAS_PARADA").ToString())
                Lista_Parada.Items.Add(ls)
            End While
            objDataReader.Close()
            desconectar()
        Catch ex As Exception
            MessageBox.Show("Ocorreu um erro: " & ex.Message)
        End Try
    End Sub



    Private Sub txt_data_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_data.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
        If e.KeyChar = ChrW(Keys.Enter) Then
            Consulta()
        End If
    End Sub

    Private Sub txt_os_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_os.KeyPress
        Dim deferenca_horas As Decimal = 0
        Dim horastotais As Decimal = 0
        Dim horastotaisREAIS As Decimal = 0

        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
        If e.KeyChar = ChrW(Keys.Enter) Then
            If txt_os.Text = "" Then
                MessageBox.Show("Digite uma ordem de serviço")
                Return
            End If
            Try
                conectar()
                comandoSQL.CommandText = "SELECT OP_NOME_OP, OP_HORAS_PREV, "
                comandoSQL.CommandText &= "DATEDIFF(MINUTE, OP_DATA_INICIO, COALESCE(OP_DATA_FIM, GETDATE())) / 60.0 As HORAS_GERADAS, "
                comandoSQL.CommandText &= "(SELECT COALESCE(SUM(DATEDIFF(MINUTE, PARADA_DATA_INICIO, COALESCE(PARADA_DATA_FIM, GETDATE()))), 0) / 60.0 FROM TAB_OP_PARADAS WHERE PARADA_ID_OP = OP_ID) As SOMA_PARADA "
                comandoSQL.CommandText &= "FROM TAB_OS_RELATORIO INNER JOIN TAB_OS_OPERACAO On OS_ID = OP_ID_OS WHERE OS_ID = " & txt_os.Text & " AND  OP_DATA_INICIO Is Not NULL GROUP BY OP_NOME_OP,OP_ID, "
                comandoSQL.CommandText &= "OS_ID,OP_DATA_INICIO, OP_DATA_FIM, OP_HORAS_PREV ORDER BY OP_ID"
                objDataReader = comandoSQL.ExecuteReader
                DataGridView1.Rows.Clear()
                If objDataReader.HasRows Then
                    While objDataReader.Read()
                        Dim horas_geradas As Decimal = If(objDataReader("HORAS_GERADAS") IsNot DBNull.Value, Convert.ToDecimal(objDataReader("HORAS_GERADAS")), 0)
                        Dim horas_parada As Decimal = If(objDataReader("SOMA_PARADA") IsNot DBNull.Value, Convert.ToDecimal(objDataReader("SOMA_PARADA")), 0)
                        deferenca_horas = Math.Round(horas_geradas - horas_parada, 2)
                        Dim row As String() = New String() {
                        objDataReader("OP_NOME_OP").ToString(),
                        objDataReader("OP_HORAS_PREV").ToString(),
                        deferenca_horas
                }
                        DataGridView1.Rows.Add(row)
                        horastotais = horastotais + objDataReader("OP_HORAS_PREV").ToString()
                        horastotaisREAIS = horastotaisREAIS + deferenca_horas
                    End While
                    Dim total As String() = New String() {
                    "Total",
                    horastotais,
                    horastotaisREAIS
                    }
                    DataGridView1.Rows.Add(total)
                Else
                    MessageBox.Show("Ordem de serviço não existe ou não tem operações")
                End If
                objDataReader.Close()
                desconectar()
            Catch ex As Exception
                MessageBox.Show("Ocorreu um erro: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub txt_op_SelectedIndexChanged(sender As Object, e As EventArgs) Handles txt_op.SelectedIndexChanged
        Consulta()
    End Sub

    Private Sub txt_op_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_op.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            Consulta()
        End If
    End Sub

    Private Sub cbo_maq_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo_maq.SelectedIndexChanged
        Consulta()
    End Sub

    Private Sub cbo_maq_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cbo_maq.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            Consulta()
        End If
    End Sub

    Private Sub ExportToExcel()
        If Lista_OS.Items.Count = 0 AndAlso Lista_Parada.Items.Count = 0 Then
            MessageBox.Show("Sem dados para exportar", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        Enabled = False
        MessageBox.Show("Gerando documento, por favor aguarde...", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Dim excelApp As Excel.Application = Nothing
        Dim workbook As Excel.Workbook = Nothing
        Dim worksheetOS As Excel.Worksheet = Nothing
        Dim worksheetParada As Excel.Worksheet = Nothing

        Try
            excelApp = New Excel.Application()
            excelApp.Visible = False
            workbook = excelApp.Workbooks.Add()

            worksheetParada = workbook.Worksheets.Add()
            worksheetParada.Name = "Paradas"

            Dim headerIndex As Integer = 1
            For Each column As ColumnHeader In Lista_Parada.Columns
                worksheetParada.Cells(1, headerIndex) = column.Text
                worksheetParada.Cells(1, headerIndex).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                headerIndex += 1
            Next

            Dim rowIndex As Integer = 2
            For Each item As ListViewItem In Lista_Parada.Items
                Dim columnIndex As Integer = 1
                For Each subItem As ListViewItem.ListViewSubItem In item.SubItems
                    worksheetParada.Cells(rowIndex, columnIndex) = subItem.Text
                    worksheetParada.Cells(rowIndex, columnIndex).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    columnIndex += 1
                Next
                rowIndex += 1
            Next

            worksheetParada.Columns.AutoFit()

            worksheetOS = workbook.Worksheets.Add()
            worksheetOS.Name = "Operações"

            headerIndex = 1
            For Each column As ColumnHeader In Lista_OS.Columns
                worksheetOS.Cells(1, headerIndex) = column.Text
                worksheetOS.Cells(1, headerIndex).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                headerIndex += 1
            Next

            rowIndex = 2
            For Each item As ListViewItem In Lista_OS.Items
                Dim columnIndex As Integer = 1
                For Each subItem As ListViewItem.ListViewSubItem In item.SubItems
                    worksheetOS.Cells(rowIndex, columnIndex) = subItem.Text
                    worksheetOS.Cells(rowIndex, columnIndex).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    columnIndex += 1
                Next
                rowIndex += 1
            Next

            worksheetOS.Columns.AutoFit()

            Dim filePath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Operações_" & DateTime.Now.ToString("dd-MM-yyyy HH-mm") & ".xlsx")
            workbook.SaveAs(filePath)
            MessageBox.Show("Documento gerado na pasta Documentos", "Documento", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Ocorreu um erro: " & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If Not worksheetOS Is Nothing Then Marshal.ReleaseComObject(worksheetOS)
            If Not worksheetParada Is Nothing Then Marshal.ReleaseComObject(worksheetParada)
            If Not workbook Is Nothing Then workbook.Close(SaveChanges:=True)
            If Not excelApp Is Nothing Then
                excelApp.Quit()
                Marshal.ReleaseComObject(excelApp)
            End If
        End Try
        Enabled = True
    End Sub



    Private Sub txt_fer_TextChanged(sender As Object, e As EventArgs) Handles txt_fer.TextChanged
        Consulta()
    End Sub

    Private Sub txt_fun_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
        If e.KeyChar = ChrW(Keys.Enter) Then
            Consulta()
        End If
    End Sub

    Private Sub txt_os_TextChanged(sender As Object, e As EventArgs) Handles txt_os.TextChanged
        Consulta()
    End Sub

    Private Sub txt_data_fim_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_data_fim.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
        If e.KeyChar = ChrW(Keys.Enter) Then
            Consulta()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btn_os.Click
        If btn_os.Text = "Criar OS" Then
            Ordem.Show()
        Else
            Solicita_Ordem.Show()
        End If
        Close()
    End Sub

    Private Sub btn_servico_Click(sender As Object, e As EventArgs) Handles btn_servico.Click
        Servico.Show()
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

    Private Sub btn_excel_Click(sender As Object, e As EventArgs) Handles btn_excel.Click
        ExportToExcel()
    End Sub

    Private Sub btn_limpar_Click(sender As Object, e As EventArgs) Handles btn_limpar.Click
        txt_op.Text = ""
        txt_fer.Text = ""
        txt_data.Text = ""
        txt_os.Text = ""
        cbo_maq.Text = ""
        Consulta()
        DataGridView1.Rows.Clear()
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