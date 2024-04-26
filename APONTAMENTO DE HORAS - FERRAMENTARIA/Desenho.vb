Public Class Desenho

    Private Sub Desenho_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Dim ImageUrl As String = Nothing
    Private Sub btn_procurar_Click(sender As Object, e As EventArgs) Handles btn_procurar.Click
        Dim ofd As New OpenFileDialog()
        ofd.Filter = "jpeg|*.jpg"

        If ofd.ShowDialog() = DialogResult.OK Then
            ImageUrl = ofd.FileName
            PictureBox.Image = Image.FromFile(ofd.FileName)

            txt_caminhadaimagem.Text = ofd.FileName
        End If
    End Sub


    Private Sub btn_imprimir_Click(sender As Object, e As EventArgs) Handles btn_imprimir.Click
        Dim dt = Ordem.ObterDados()
        Dim caixa As Integer
        If Ordem.txt_caixa.Text = "" Then
            caixa = 0
        Else
            caixa = Ordem.txt_caixa.Text
        End If
        Using frm = New TelaImpressao(Ordem.COD_FORM, Ordem.txt_desc.Text, Ordem.txt_pos.Text, Ordem.txt_fer.Text, Ordem.txt_resp.Text, Ordem.txt_dt_ini.Text, Ordem.txt_dt_ent.Text, Ordem.txt_conta.Text, Ordem.txt_sub_conta.Text, Ordem.txt_secao.Text, Ordem.txt_qtd.Text, Ordem.txt_tp.Text, caixa, dt, txt_caminhadaimagem.Text)
            frm.ShowDialog()
        End Using
        Close()
    End Sub
End Class



