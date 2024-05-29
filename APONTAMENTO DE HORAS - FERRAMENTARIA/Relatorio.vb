Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Microsoft.ReportingServices.Rendering.ExcelRenderer
Imports OfficeOpenXml

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
        Try
            conectar()
            comandoSQL.CommandText = "WITH ParadasNumeradas AS ( "
            comandoSQL.CommandText &= "SELECT  o.OP_ID, o.OP_NOME_OP, o.OP_FUNCIONARIO, o.OP_DATA_INICIO, o.OP_DATA_FIM, o.OP_MAQUINA, o.OP_HORAS_PREV, o.OP_QUANTIDADE, o.OP_QUANTIDADE_FINAL, p1.PARADA_DATA_INICIO, p1.PARADA_DATA_FIM, "
            comandoSQL.CommandText &= "(SELECT COUNT(*) FROM TAB_OP_PARADAS p2 WHERE p2.PARADA_ID_OP = p1.PARADA_ID_OP AND p2.PARADA_DATA_INICIO < p1.PARADA_DATA_INICIO) + 1 AS RN "
            comandoSQL.CommandText &= "FROM TAB_OS_OPERACAO o LEFT JOIN TAB_OP_PARADAS p1 ON p1.PARADA_ID_OP = o.OP_ID), "
            comandoSQL.CommandText &= "Intervalos AS (SELECT t1.OP_ID, t1.OP_NOME_OP, t1.OP_FUNCIONARIO, t1.OP_DATA_INICIO, t1.OP_DATA_FIM, t1.OP_MAQUINA, "
            comandoSQL.CommandText &= "t1.OP_HORAS_PREV, t1.OP_QUANTIDADE, t1.OP_QUANTIDADE_FINAL, COALESCE(t2.PARADA_DATA_FIM, t1.OP_DATA_INICIO) AS INICIO,COALESCE(t1.PARADA_DATA_INICIO, t1.OP_DATA_FIM) AS FIM "
            comandoSQL.CommandText &= "FROM ParadasNumeradas t1 LEFT JOIN ParadasNumeradas t2 ON t1.OP_ID = t2.OP_ID AND t1.RN = t2.RN + 1) "
            comandoSQL.CommandText &= "SELECT r.OS_ID, r.OS_FERRAMENTA, r.OS_POSICAO, r.OS_SECAO, r.OS_PROJETO, r.OS_CONTA, r.OS_SUB_CONTA, r.OS_REG_RESPONSAVEL, r.OS_TIPO, r.OS_DATA_INICIO, r.OS_QUANTIDADE, r.OS_QUANTIDADE_FINAL, r.OS_DATA_FIM,"
            comandoSQL.CommandText &= "i.OP_ID, i.OP_NOME_OP, i.OP_FUNCIONARIO, i.OP_DATA_INICIO, i.OP_DATA_FIM, i.OP_MAQUINA, i.OP_HORAS_PREV, i.OP_QUANTIDADE, i.OP_QUANTIDADE_FINAL, i.INICIO, i.FIM, DATEDIFF(MINUTE, i.INICIO, i.FIM) / 60.0 AS TEMPO "
            comandoSQL.CommandText &= "FROM Intervalos i INNER JOIN TAB_OS_OPERACAO o ON i.OP_ID = o.OP_ID INNER JOIN TAB_OS_RELATORIO r ON o.OP_ID_OS = r.OS_ID "
            comandoSQL.CommandText &= "WHERE i.OP_DATA_INICIO IS NOT NULL "
            comandoSQL.CommandText &= "AND ('" & txt_fer.Text & "' = '' OR r.OS_FERRAMENTA LIKE '%" & txt_fer.Text & "%' OR r.OS_POSICAO LIKE '%" & txt_fer.Text & "%') "
            comandoSQL.CommandText &= "AND ('" & txt_data.Text & "' = '  /  /' OR CONVERT(DATE, i.INICIO, 103) >= CONVERT(DATE, '" & txt_data.Text & "', 103)) "
            comandoSQL.CommandText &= "AND ('" & txt_data_fim.Text & "' = '  /  /' OR CONVERT(DATE, i.FIM, 103) <= CONVERT(DATE, '" & txt_data_fim.Text & "', 103)) "
            comandoSQL.CommandText &= "AND ('" & txt_op.Text & "' = '' OR i.OP_NOME_OP = '" & txt_op.Text & "')  AND ('" & cbo_maq.Text & "' = '' OR i.OP_MAQUINA = '" & cbo_maq.Text & "') "
            comandoSQL.CommandText &= "AND ('" & txt_os.Text & "' = '' OR r.OS_ID = '" & txt_os.Text & "') "
            comandoSQL.CommandText &= "ORDER BY r.OS_ID, i.OP_ID;"
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
            Lista_OS.Columns.Add("RESPONSÁVEL", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("TIPO", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("QUANTIDADE", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("QUANTIDADE FINAL", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("INÍCIO/ORDEM", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("FIM/ORDEM", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("OPERAÇÃO", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("NOME", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("FUNCIONÁRIO", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("MAQUINA", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("INÍCIO/OPERAÇÃO", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("FIM/OPERAÇÃO", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("HORAS PROGRAMADAS", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("QUANTIDADE", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("QUANTIDADE FINAL", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("INÍCIO/APONTAMENTO", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("FIM/APONTAMENTO", 60, HorizontalAlignment.Center)
            Lista_OS.Columns.Add("HORAS REAIS", 60, HorizontalAlignment.Center)

            While objDataReader.Read()
                Dim ls As New ListViewItem(objDataReader("OS_ID").ToString())
                ls.SubItems.Add(objDataReader("OS_FERRAMENTA").ToString())
                ls.SubItems.Add(objDataReader("OS_POSICAO").ToString())
                ls.SubItems.Add(objDataReader("OS_PROJETO").ToString())
                ls.SubItems.Add(objDataReader("OS_SECAO").ToString())
                ls.SubItems.Add(objDataReader("OS_CONTA").ToString())
                ls.SubItems.Add(objDataReader("OS_SUB_CONTA").ToString())
                ls.SubItems.Add(objDataReader("OS_REG_RESPONSAVEL").ToString())
                ls.SubItems.Add(objDataReader("OS_TIPO").ToString())
                ls.SubItems.Add(objDataReader("OS_QUANTIDADE").ToString())
                ls.SubItems.Add(objDataReader("OS_QUANTIDADE_FINAL").ToString())
                ls.SubItems.Add(If(objDataReader("OS_DATA_INICIO") IsNot DBNull.Value, Convert.ToDateTime(objDataReader("OS_DATA_INICIO")).ToString("dd/MM/yyyy HH:mm"), Nothing))
                ls.SubItems.Add(If(objDataReader("OS_DATA_FIM") IsNot DBNull.Value, Convert.ToDateTime(objDataReader("OS_DATA_FIM")).ToString("dd/MM/yyyy HH:mm"), Nothing))
                ls.SubItems.Add(objDataReader("OP_ID").ToString())
                ls.SubItems.Add(objDataReader("OP_NOME_OP").ToString())
                ls.SubItems.Add(objDataReader("OP_FUNCIONARIO").ToString())
                ls.SubItems.Add(objDataReader("OP_MAQUINA").ToString())
                ls.SubItems.Add(If(objDataReader("OP_DATA_INICIO") IsNot DBNull.Value, Convert.ToDateTime(objDataReader("OP_DATA_INICIO")).ToString("dd/MM/yyyy HH:mm"), Nothing))
                ls.SubItems.Add(If(objDataReader("OP_DATA_FIM") IsNot DBNull.Value, Convert.ToDateTime(objDataReader("OP_DATA_FIM")).ToString("dd/MM/yyyy HH:mm"), Nothing))
                ls.SubItems.Add(objDataReader("OP_HORAS_PREV").ToString())
                ls.SubItems.Add(objDataReader("OP_QUANTIDADE").ToString())
                ls.SubItems.Add(objDataReader("OP_QUANTIDADE_FINAL").ToString())
                ls.SubItems.Add(If(objDataReader("INICIO") IsNot DBNull.Value, Convert.ToDateTime(objDataReader("INICIO")).ToString("dd/MM/yyyy HH:mm"), Nothing))
                ls.SubItems.Add(If(objDataReader("FIM") IsNot DBNull.Value, Convert.ToDateTime(objDataReader("FIM")).ToString("dd/MM/yyyy HH:mm"), Nothing))
                ls.SubItems.Add(If(objDataReader("TEMPO") IsNot DBNull.Value, Convert.ToDecimal(objDataReader("TEMPO")), Nothing))
                Lista_OS.Items.Add(ls)
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
        If Lista_OS.Items.Count = 0 Then
            MessageBox.Show("Sem dados para exportar", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Enabled = False
        MessageBox.Show("Gerando documento, por favor aguarde...", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Try
            ' Definir o contexto de licença do EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial

            ' Definir o caminho do arquivo Excel
            Dim excelPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Operações " & DateTime.Now.ToString("dd-MM-yyyy HH-mm") & ".xlsx")

            ' Criar um novo pacote Excel
            Using package As New ExcelPackage()
                ' Adicionar uma nova worksheet
                Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Dados Exportados")

                ' Adicionar os cabeçalhos das colunas
                For col As Integer = 0 To Lista_OS.Columns.Count - 1
                    worksheet.Cells(1, col + 1).Value = Lista_OS.Columns(col).Text
                Next

                ' Adicionar os dados das linhas
                For row As Integer = 0 To Lista_OS.Items.Count - 1
                    For col As Integer = 0 To Lista_OS.Columns.Count - 1
                        worksheet.Cells(row + 2, col + 1).Value = Lista_OS.Items(row).SubItems(col).Text
                    Next
                Next

                ' Salvar o arquivo Excel
                Dim fileInfo As New FileInfo(excelPath)
                package.SaveAs(fileInfo)
            End Using

            MessageBox.Show("Documento gerado na pasta Documentos", "Documento", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Ocorreu um erro: " & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Enabled = True
        End Try
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
