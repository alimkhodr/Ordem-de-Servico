Imports System.Data.SqlClient
Module Metodo
    Public Verifica As Integer
    Public User As Integer
    Public UserName As String
    Public conexaobd As New SqlConnection()
    Public objDataReader As SqlDataReader
    Public comandoSQL As SqlCommand
    Public id_maquina, id_OS As Integer
    Public funtipo As Integer
    Public Sub conectar()
        Try
            conexaobd.ConnectionString = "server=BRSPJAM-AP07;database=SGM_ONE;uid=sa;pwd=P@ssw0rd"
            comandoSQL = conexaobd.CreateCommand
            conexaobd.Open()

        Catch ex As Exception
            conexaobd.Close()
            conexaobd.ConnectionString = "server=BRSPJAM-AP07;database=SGM_ONE;uid=sa;pwd=P@ssw0rd"
            comandoSQL = conexaobd.CreateCommand
            conexaobd.Open()
        End Try

    End Sub
    Public Sub desconectar()
        Try
            conexaobd.Close()
            objDataReader.Close()
        Catch ex As Exception

        End Try

    End Sub

    Public Sub permissoes()
        Dim names() As String = UserName.Split(" "c)
        If names.Length >= 2 Then
            Dim initials As String = names(0)(0) & names(names.Length - 1)(0)
            Tela_Inicial.btn_user.Text = initials
            Ordem.btn_user.Text = initials
            Solicita_Ordem.btn_user.Text = initials
            Relatorio.btn_user.Text = initials
        End If

        If funtipo = 1 Then
            Ordem.btn_os.Text = "Criar OS"
            Ordem.btn_servico.Visible = True
            Ordem.btn_os.Visible = True
            Ordem.btn_maq.Visible = True
            Ordem.btn_relatorio.Visible = True

            Tela_Inicial.btn_os.Text = "Criar OS"
            Tela_Inicial.btn_servico.Visible = True
            Tela_Inicial.btn_os.Visible = True
            Tela_Inicial.btn_maq.Visible = True
            Tela_Inicial.btn_relatorio.Visible = True

            Relatorio.btn_os.Text = "Criar OS"
            Relatorio.btn_servico.Visible = True
            Relatorio.btn_os.Visible = True
            Relatorio.btn_maq.Visible = True
            Relatorio.btn_relatorio.Visible = True

            Solicita_Ordem.btn_os.Text = "Criar OS"
            Solicita_Ordem.btn_os.Visible = True
        ElseIf funtipo = 2 Then
            Ordem.btn_servico.Visible = True
            Ordem.btn_os.Visible = False
            Ordem.btn_maq.Visible = False
            Ordem.btn_relatorio.Visible = False

            Tela_Inicial.btn_servico.Visible = True
            Tela_Inicial.btn_os.Visible = False
            Tela_Inicial.btn_maq.Visible = False
            Tela_Inicial.btn_relatorio.Visible = False

            Relatorio.btn_servico.Visible = True
            Relatorio.btn_excel.Visible = False
            Relatorio.btn_maq.Visible = False
            Relatorio.btn_relatorio.Visible = False

            Solicita_Ordem.btn_os.Visible = False
        ElseIf funtipo = 3 Then
            Ordem.btn_servico.Visible = False
            Ordem.btn_os.Visible = True
            Ordem.btn_maq.Visible = False
            Ordem.btn_relatorio.Visible = False
            Ordem.btn_os.Text = "Solicitar OS"

            Tela_Inicial.btn_servico.Visible = False
            Tela_Inicial.btn_os.Visible = True
            Tela_Inicial.btn_maq.Visible = False
            Tela_Inicial.btn_relatorio.Visible = False
            Tela_Inicial.btn_os.Text = "Solicitar OS"

            Relatorio.btn_servico.Visible = False
            Relatorio.btn_excel.Visible = True
            Relatorio.btn_maq.Visible = False
            Relatorio.btn_relatorio.Visible = False
            Relatorio.btn_excel.Text = "Solicitar OS"

            Solicita_Ordem.btn_os.Visible = True
            Solicita_Ordem.btn_os.Text = "Solicitar OS"
        End If
    End Sub

End Module
