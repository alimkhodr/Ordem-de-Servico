Imports Microsoft.Reporting.WinForms

Public Class TelaImpressao

    Dim dt = New DataTable()

    Sub New(ByVal os As Integer, ByVal desc As String, ByVal pos As String, ByVal fer As String, ByVal resp As String, ByVal dti As DateTime, ByVal dtf As DateTime, ByVal conta As Integer, ByVal sub_conta As Integer, ByVal secao As Integer, ByVal qtd As Integer, ByVal tp As String, ByVal caixa As Integer, dt As DataTable, ByVal imagem As String)
        InitializeComponent()
        Me.dt = dt

        ReportViewer1.LocalReport.DataSources.Clear()
        ReportViewer1.LocalReport.ReportEmbeddedResource = "Ordem_de_Servico.Impressao.rdlc"
        Dim p As ReportParameter() = New ReportParameter(13) {}
        p(0) = New ReportParameter("OS", os)
        p(1) = New ReportParameter("OS_desc", desc)
        p(2) = New ReportParameter("OS_pos", pos)
        p(3) = New ReportParameter("OS_ferr", fer)
        p(4) = New ReportParameter("OS_resp", resp)
        p(5) = New ReportParameter("OS_dti", dti)
        p(6) = New ReportParameter("OS_dtf", dtf)
        p(7) = New ReportParameter("OS_conta", conta)
        p(8) = New ReportParameter("OS_sub", sub_conta)
        p(9) = New ReportParameter("OS_secao", secao)
        p(10) = New ReportParameter("OS_qtd", qtd)
        p(11) = New ReportParameter("OS_tp", tp)
        If caixa = "0" Then
            p(12) = New ReportParameter("OS_caixa", "")
        Else
            p(12) = New ReportParameter("OS_caixa", caixa)
        End If
        If imagem = "imagem.jpg" Then
            p(13) = New ReportParameter("rpt_img", "", False)
        Else
            p(13) = New ReportParameter("rpt_img", "File:\" + imagem, True)
        End If
        ReportViewer1.LocalReport.EnableExternalImages = True
        ReportViewer1.LocalReport.SetParameters(p)
        ReportViewer1.LocalReport.Refresh()
        ReportViewer1.RefreshReport()
    End Sub



    Private Sub TelaImpressao_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.ReportViewer1.LocalReport.DataSources.Clear()
        Me.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("DataSet1", dt))
        Me.ReportViewer1.RefreshReport()
    End Sub
End Class