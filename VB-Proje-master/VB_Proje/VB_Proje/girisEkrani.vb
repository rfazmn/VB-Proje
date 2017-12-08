Imports System.Data.Sql
Imports System.Data.SqlClient

Public Class kullaniciGiris_ekrani

    Private baglanti As New SqlConnection


    Private Sub btn_giris_Click(sender As Object, e As EventArgs) Handles btn_giris.Click
        Dim sqlQuery As String
        sqlQuery = "Select KullaniciAdi,Sifre,Yoneticisi From Personeller Where KullaniciAdi = @kullaniciadi And Sifre = @sifre"
        Dim command As New SqlCommand(sqlQuery, baglanti)

        command.Parameters.Add("@kullaniciadi", SqlDbType.NVarChar).Value = tb_kullaniciAdi.Text
        command.Parameters.Add("@sifre", SqlDbType.NVarChar).Value = tb_sifre.Text


        Dim adapter As New SqlDataAdapter(command)
        Dim table As New DataTable()
        adapter.Fill(table)


        Dim row As DataRow

        If table.Rows.Count > 0 Then
            row = table.Rows(0)
        End If


        If table.Rows.Count <= 0 Then
            MessageBox.Show("Kullanıcı adı veya şifre hatalı!")

        Else
            If Not row.IsNull(2) Then
                personelAnasafya.Show()
                If cb_hatirla.Checked Then

                Else
                    GirisTemizle()
                End If
                Me.Hide()

            Else
                yoneticiAnasayfa.Show()
                If cb_hatirla.Checked Then

                Else
                    GirisTemizle()
                End If
                Me.Hide()
            End If
        End If

        'Kullanıcı bilgisini kaydetme
        'If cb_hatirla.Checked Then
        '    My.Settings.KullaniciAdi = tb_kullaniciAdi.Text
        '    My.Settings.Sifre = tb_sifre.Text
        '    My.Settings.BeniHatirla = True
        'End If

    End Sub

    Private Sub kullaniciGiris_ekrani_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        yoneticiAnasayfa.Show()


        Try
            baglanti = New SqlConnection("Server=FENERBAHCE;Database=ProjeDb;User Id=Deneme;Password=123456;")
            baglanti.Open()

        Catch ex As SqlException
            MessageBox.Show("Hata" & ex.Message)
        End Try

        'Kullanıcı bilgisini getirme
        'If My.Settings.BeniHatirla Then
        '    tb_kullaniciAdi.Text = My.Settings.KullaniciAdi
        '    tb_sifre.Text = My.Settings.Sifre
        'End If
    End Sub

    Sub GirisTemizle()
        tb_kullaniciAdi.Text = ""
        tb_sifre.Text = ""
    End Sub
End Class
