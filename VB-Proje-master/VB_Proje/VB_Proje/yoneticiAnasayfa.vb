Imports System.Data.SqlClient

Public Class yoneticiAnasayfa
    Dim kisiler As String = ""
    Dim secilenKisiler() As String
    Dim eklencekMi As Boolean = True

    Dim baglanti As SqlConnection
    Dim sqlQuery As String
    Dim command As New SqlCommand
    Dim adapter As New SqlDataAdapter
    Dim table As New DataTable()
    Dim gorevDetay(10) As String

    Structure IdBilgisi
        Dim GorevTurID() As Integer
        Dim GorevYerID() As Integer
        Dim GorevBirimID() As Integer
        Dim personelID() As Integer
    End Structure
    Dim ID_Bilgileri As New IdBilgisi

    Private Sub yoneticiAnasayfa_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AlanBilgieriniGetir()
    End Sub


    Private Sub btn_bitenGorevler_Click(sender As Object, e As EventArgs) Handles btn_bitenGorevler.Click
        bitenGorevler.Show()
    End Sub

    Private Sub btn_personelBilgisi_Click(sender As Object, e As EventArgs) Handles btn_personelBilgisi.Click
        personelBilgileri.Show()
    End Sub

    Private Sub btn_cikis_Click(sender As Object, e As EventArgs) Handles btn_cikis.Click
        Me.Close()
    End Sub

    Private Sub btn_gorevEkle_Click(sender As Object, e As EventArgs) Handles btn_gorevEkle.Click
        eklencekMi = True
        For Each kisi In cbl_personeller.CheckedItems
            kisiler += kisi + " "
        Next

        If dtp_baslangic.Value > dtp_bitis.Value Then
            MessageBox.Show("Lütfen Geçerli Bir Tarih Girin.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            eklencekMi = False
        End If

        If Kontrol() Then
            Try
                SeciliKisiler()
                Dim deneme As String = "Birim : " + cb_birim.SelectedItem.ToString + Environment.NewLine +
                    "Görev Yeri : " + cb_gorevYeri.SelectedItem.ToString + Environment.NewLine + "Görev Türü : " + cb_gorevTuru.SelectedItem.ToString + Environment.NewLine +
                "Personeler : " + kisiler +
                Environment.NewLine + Environment.NewLine + tb_aciklama.Text + Environment.NewLine + "Başlangıç Tarihi : " + dtp_baslangic.Value.ToShortDateString + Environment.NewLine + "Bitiş Tarihi : " + dtp_bitis.Value.ToShortDateString
            Catch ex As Exception
                MessageBox.Show("Lütfen Tüm Alanları Doldurun.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                eklencekMi = False
            End Try

            If eklencekMi Then
                GorevEkle()
                MessageBox.Show("Görev ekleme başarılı", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information)
                AlanlarıYenile()
            End If
        End If
    End Sub

    Sub AlanlarıYenile()
        lb_gorevler.Items.Clear()
        tb_gorevAdi.Clear()
        tb_aciklama.Clear()
        cb_gorevTuru.SelectedItem = vbNull
        cb_birim.SelectedItem = vbNull
        cb_gorevYeri.SelectedItem = vbNull
        For Each x In cbl_personeller.CheckedIndices
            cbl_personeller.SetItemCheckState(x, CheckState.Unchecked)
        Next
        GorevGetir()
        AciklamaGetir()
    End Sub

    Sub GorevEkle()

        For i = 0 To ID_Bilgileri.personelID.Count - 1
            If ID_Bilgileri.personelID(i) = 0 Then
                Continue For
            Else
                With command
                    command.Parameters.Clear()

                    command.CommandText = "INSERT INTO Gorevler (GorevAdı, GorevAciklamasi, GorevTuru, GorevBirimi, BaslangicTarihi, TahminiBitisTarihi,AtananPersonel,GorevYeri) 
                                   VALUES (@ad,@aciklama, @tur, @birim,@bTarih, @tBTarih,@personel,@yeri)"

                    .Parameters.AddWithValue("@ad", tb_gorevAdi.Text)
                    .Parameters.AddWithValue("@aciklama", tb_aciklama.Text)
                    .Parameters.AddWithValue("@tur", ID_Bilgileri.GorevTurID(cb_gorevTuru.SelectedIndex))
                    .Parameters.AddWithValue("@yeri", ID_Bilgileri.GorevYerID(cb_gorevYeri.SelectedIndex))
                    .Parameters.AddWithValue("@birim", ID_Bilgileri.GorevBirimID(cb_birim.SelectedIndex))
                    .Parameters.AddWithValue("@bTarih", dtp_baslangic.Value)
                    .Parameters.AddWithValue("@tBTarih", dtp_bitis.Value)
                    .Parameters.AddWithValue("@personel", ID_Bilgileri.personelID(i))
                    command.ExecuteNonQuery()
                End With
            End If

        Next


    End Sub

    Function Kontrol() As Boolean
        If kisiler <> "" And tb_aciklama.Text <> "" And tb_gorevAdi.Text <> "" Then
            Return True
        End If
        MessageBox.Show("Lütfen Tüm Alanları Doldurun.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Return False
    End Function

    Sub GorevGetir()
        sqlQuery = "Select GorevAdı from Gorevler"

        command.CommandText = sqlQuery
        adapter.SelectCommand = command
        table.Reset()
        adapter.Fill(table)

        For i = 0 To table.Rows.Count - 1
            lb_gorevler.Items.Add(table.Rows(i).Item(0).ToString)
        Next

    End Sub

    Sub AciklamaGetir()
        sqlQuery = "Select GorevAciklamasi from Gorevler"

        command.CommandText = sqlQuery
        adapter.SelectCommand = command
        table.Reset()
        adapter.Fill(table)

        If table.Rows.Count > 10 Then
            ReDim Preserve gorevDetay(table.Rows.Count)
        End If

        For i = 0 To table.Rows.Count - 1
            gorevDetay(i) = table.Rows(i).Item(0).ToString
        Next


    End Sub

    Sub SeciliKisiler()


        Dim personelIdTemp As Integer

        For i = 0 To kisiler.Length
            secilenKisiler = kisiler.Split(" ")
        Next

        ReDim Preserve ID_Bilgileri.personelID(secilenKisiler.Length)


        For i = 0 To secilenKisiler.Length - 2 Step 2

            kisiler = secilenKisiler(i)

            sqlQuery = "Select KullaniciID from Personeller where Ad='" + kisiler + "'"

            command.CommandText = sqlQuery
            adapter.SelectCommand = command
            table.Reset()
            adapter.Fill(table)

            Try
                personelIdTemp = table.Rows(0).Item(0)
                ID_Bilgileri.personelID(i) = personelIdTemp
            Catch ex As Exception
            End Try
        Next

    End Sub

    Sub AlanBilgieriniGetir()
        Try
            baglanti = New SqlConnection("Server=FENERBAHCE;Database=ProjeDb;User Id=Deneme;Password=123456;")
            baglanti.Open()

            sqlQuery = "Select TurID, TurAdi from GorevTurleri"

            command.Connection = baglanti
            command.CommandText = sqlQuery
            adapter.SelectCommand = command
            adapter.Fill(table)

            ReDim Preserve ID_Bilgileri.GorevTurID(table.Rows.Count)

            For i = 0 To table.Rows.Count - 1
                ID_Bilgileri.GorevTurID(i) = table.Rows(i).Item(0).ToString
                cb_gorevTuru.Items.Add(table.Rows(i).Item(1).ToString)
            Next

            sqlQuery = "Select BirimID, BirimAdi from Birimler"

            command.CommandText = sqlQuery
            adapter.SelectCommand = command
            table.Reset()
            adapter.Fill(table)

            ReDim Preserve ID_Bilgileri.GorevBirimID(table.Rows.Count)

            For i = 0 To table.Rows.Count - 1
                ID_Bilgileri.GorevBirimID(i) = table.Rows(i).Item(0).ToString
                cb_birim.Items.Add(table.Rows(i).Item(1).ToString)
            Next

            sqlQuery = "Select NesneID, NesneAdi from GorevYerleri"

            command.CommandText = sqlQuery
            adapter.SelectCommand = command
            table.Reset()
            adapter.Fill(table)

            ReDim Preserve ID_Bilgileri.GorevYerID(table.Rows.Count)

            For i = 0 To table.Rows.Count - 1
                ID_Bilgileri.GorevYerID(i) = table.Rows(i).Item(0).ToString
                cb_gorevYeri.Items.Add(table.Rows(i).Item(1).ToString)
            Next

            sqlQuery = "Select Ad, Soyad from Personeller"

            command.CommandText = sqlQuery
            adapter.SelectCommand = command
            table.Reset()
            adapter.Fill(table)

            For i = 0 To table.Rows.Count - 1
                cbl_personeller.Items.Add(table.Rows(i).Item(0).ToString & " " & table.Rows(i).Item(1).ToString)
            Next

            GorevGetir()
            AciklamaGetir()

        Catch ex As SqlException
            MessageBox.Show("Hata" & ex.Message)
        End Try
    End Sub

    Private Sub lb_gorevler_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lb_gorevler.SelectedIndexChanged
        lbl_gorevAyrinti.Text = gorevDetay(lb_gorevler.SelectedIndex)
    End Sub

    Private Sub yenileClick(sender As Object, e As EventArgs) Handles btn_yenile.Click
        AlanlarıYenile()
    End Sub
End Class