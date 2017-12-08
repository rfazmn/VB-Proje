<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class personelAnasafya
    Inherits System.Windows.Forms.Form

    'Form, bileşen listesini temizlemeyi bırakmayı geçersiz kılar.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows Form Tasarımcısı tarafından gerektirilir
    Private components As System.ComponentModel.IContainer

    'NOT: Aşağıdaki yordam Windows Form Tasarımcısı için gereklidir
    'Windows Form Tasarımcısı kullanılarak değiştirilebilir.  
    'Kod düzenleyicisini kullanarak değiştirmeyin.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.gb_bilgiler = New System.Windows.Forms.GroupBox()
        Me.lbl_yoneticiTel = New System.Windows.Forms.Label()
        Me.lbl_tel = New System.Windows.Forms.Label()
        Me.lbl_birim = New System.Windows.Forms.Label()
        Me.lbl_soyad = New System.Windows.Forms.Label()
        Me.lbl_ad = New System.Windows.Forms.Label()
        Me.gb_gorevler = New System.Windows.Forms.GroupBox()
        Me.btn_cikis = New System.Windows.Forms.Button()
        Me.btn_gorevTamamla = New System.Windows.Forms.Button()
        Me.lbl_gorevAciklama = New System.Windows.Forms.Label()
        Me.lbl_gorevNotu = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tb_gorevNotu = New System.Windows.Forms.TextBox()
        Me.lb_gorevler = New System.Windows.Forms.ListBox()
        Me.gb_bilgiler.SuspendLayout()
        Me.gb_gorevler.SuspendLayout()
        Me.SuspendLayout()
        '
        'gb_bilgiler
        '
        Me.gb_bilgiler.Controls.Add(Me.lbl_yoneticiTel)
        Me.gb_bilgiler.Controls.Add(Me.lbl_tel)
        Me.gb_bilgiler.Controls.Add(Me.lbl_birim)
        Me.gb_bilgiler.Controls.Add(Me.lbl_soyad)
        Me.gb_bilgiler.Controls.Add(Me.lbl_ad)
        Me.gb_bilgiler.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.gb_bilgiler.Location = New System.Drawing.Point(12, 12)
        Me.gb_bilgiler.Name = "gb_bilgiler"
        Me.gb_bilgiler.Size = New System.Drawing.Size(720, 72)
        Me.gb_bilgiler.TabIndex = 0
        Me.gb_bilgiler.TabStop = False
        Me.gb_bilgiler.Text = "Bilgilerim"
        '
        'lbl_yoneticiTel
        '
        Me.lbl_yoneticiTel.AutoSize = True
        Me.lbl_yoneticiTel.Location = New System.Drawing.Point(455, 44)
        Me.lbl_yoneticiTel.Name = "lbl_yoneticiTel"
        Me.lbl_yoneticiTel.Size = New System.Drawing.Size(111, 16)
        Me.lbl_yoneticiTel.TabIndex = 3
        Me.lbl_yoneticiTel.Text = "Yönetici Telefon :"
        '
        'lbl_tel
        '
        Me.lbl_tel.AutoSize = True
        Me.lbl_tel.Location = New System.Drawing.Point(455, 18)
        Me.lbl_tel.Name = "lbl_tel"
        Me.lbl_tel.Size = New System.Drawing.Size(60, 16)
        Me.lbl_tel.TabIndex = 2
        Me.lbl_tel.Text = "Telefon :"
        '
        'lbl_birim
        '
        Me.lbl_birim.AutoSize = True
        Me.lbl_birim.Location = New System.Drawing.Point(325, 18)
        Me.lbl_birim.Name = "lbl_birim"
        Me.lbl_birim.Size = New System.Drawing.Size(47, 16)
        Me.lbl_birim.TabIndex = 2
        Me.lbl_birim.Text = "Birim : "
        '
        'lbl_soyad
        '
        Me.lbl_soyad.AutoSize = True
        Me.lbl_soyad.Location = New System.Drawing.Point(147, 18)
        Me.lbl_soyad.Name = "lbl_soyad"
        Me.lbl_soyad.Size = New System.Drawing.Size(60, 16)
        Me.lbl_soyad.TabIndex = 1
        Me.lbl_soyad.Text = "Soyadı : "
        '
        'lbl_ad
        '
        Me.lbl_ad.AutoSize = True
        Me.lbl_ad.Location = New System.Drawing.Point(7, 18)
        Me.lbl_ad.Name = "lbl_ad"
        Me.lbl_ad.Size = New System.Drawing.Size(31, 16)
        Me.lbl_ad.TabIndex = 0
        Me.lbl_ad.Text = "Ad :"
        '
        'gb_gorevler
        '
        Me.gb_gorevler.Controls.Add(Me.btn_cikis)
        Me.gb_gorevler.Controls.Add(Me.btn_gorevTamamla)
        Me.gb_gorevler.Controls.Add(Me.lbl_gorevAciklama)
        Me.gb_gorevler.Controls.Add(Me.lbl_gorevNotu)
        Me.gb_gorevler.Controls.Add(Me.Label1)
        Me.gb_gorevler.Controls.Add(Me.tb_gorevNotu)
        Me.gb_gorevler.Controls.Add(Me.lb_gorevler)
        Me.gb_gorevler.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.gb_gorevler.Location = New System.Drawing.Point(12, 113)
        Me.gb_gorevler.Name = "gb_gorevler"
        Me.gb_gorevler.Size = New System.Drawing.Size(720, 400)
        Me.gb_gorevler.TabIndex = 0
        Me.gb_gorevler.TabStop = False
        Me.gb_gorevler.Text = "Görevler"
        '
        'btn_cikis
        '
        Me.btn_cikis.Location = New System.Drawing.Point(10, 357)
        Me.btn_cikis.Name = "btn_cikis"
        Me.btn_cikis.Size = New System.Drawing.Size(117, 34)
        Me.btn_cikis.TabIndex = 3
        Me.btn_cikis.Text = "Çıkış Yap"
        Me.btn_cikis.UseVisualStyleBackColor = True
        '
        'btn_gorevTamamla
        '
        Me.btn_gorevTamamla.Location = New System.Drawing.Point(597, 357)
        Me.btn_gorevTamamla.Name = "btn_gorevTamamla"
        Me.btn_gorevTamamla.Size = New System.Drawing.Size(117, 34)
        Me.btn_gorevTamamla.TabIndex = 3
        Me.btn_gorevTamamla.Text = "Görevi Tamamla"
        Me.btn_gorevTamamla.UseVisualStyleBackColor = True
        '
        'lbl_gorevAciklama
        '
        Me.lbl_gorevAciklama.AutoSize = True
        Me.lbl_gorevAciklama.Location = New System.Drawing.Point(200, 18)
        Me.lbl_gorevAciklama.Name = "lbl_gorevAciklama"
        Me.lbl_gorevAciklama.Size = New System.Drawing.Size(64, 16)
        Me.lbl_gorevAciklama.TabIndex = 0
        Me.lbl_gorevAciklama.Text = "Açıklama"
        '
        'lbl_gorevNotu
        '
        Me.lbl_gorevNotu.AutoSize = True
        Me.lbl_gorevNotu.Location = New System.Drawing.Point(200, 260)
        Me.lbl_gorevNotu.Name = "lbl_gorevNotu"
        Me.lbl_gorevNotu.Size = New System.Drawing.Size(76, 16)
        Me.lbl_gorevNotu.TabIndex = 0
        Me.lbl_gorevNotu.Text = "Görev Notu"
        '
        'Label1
        '
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Location = New System.Drawing.Point(203, 37)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(511, 214)
        Me.Label1.TabIndex = 0
        '
        'tb_gorevNotu
        '
        Me.tb_gorevNotu.Location = New System.Drawing.Point(203, 279)
        Me.tb_gorevNotu.Multiline = True
        Me.tb_gorevNotu.Name = "tb_gorevNotu"
        Me.tb_gorevNotu.Size = New System.Drawing.Size(511, 66)
        Me.tb_gorevNotu.TabIndex = 2
        '
        'lb_gorevler
        '
        Me.lb_gorevler.FormattingEnabled = True
        Me.lb_gorevler.ItemHeight = 16
        Me.lb_gorevler.Items.AddRange(New Object() {"Görev 1", "Görev 2", "Görev 3", "Görev 4"})
        Me.lb_gorevler.Location = New System.Drawing.Point(6, 37)
        Me.lb_gorevler.Name = "lb_gorevler"
        Me.lb_gorevler.Size = New System.Drawing.Size(191, 308)
        Me.lb_gorevler.TabIndex = 1
        '
        'personelAnasafya
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(748, 525)
        Me.Controls.Add(Me.gb_gorevler)
        Me.Controls.Add(Me.gb_bilgiler)
        Me.MaximizeBox = False
        Me.Name = "personelAnasafya"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Personel Anasayfa"
        Me.gb_bilgiler.ResumeLayout(False)
        Me.gb_bilgiler.PerformLayout()
        Me.gb_gorevler.ResumeLayout(False)
        Me.gb_gorevler.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents gb_bilgiler As GroupBox
    Friend WithEvents gb_gorevler As GroupBox
    Friend WithEvents Label1 As Label
    Friend WithEvents tb_gorevNotu As TextBox
    Friend WithEvents lb_gorevler As ListBox
    Friend WithEvents lbl_gorevAciklama As Label
    Friend WithEvents lbl_gorevNotu As Label
    Friend WithEvents btn_gorevTamamla As Button
    Friend WithEvents lbl_birim As Label
    Friend WithEvents lbl_soyad As Label
    Friend WithEvents lbl_ad As Label
    Friend WithEvents btn_cikis As Button
    Friend WithEvents lbl_yoneticiTel As Label
    Friend WithEvents lbl_tel As Label
End Class
