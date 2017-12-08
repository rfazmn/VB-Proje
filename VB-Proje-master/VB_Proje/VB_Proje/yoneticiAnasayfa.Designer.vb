<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class yoneticiAnasayfa
    Inherits System.Windows.Forms.Form

    'Form, bileşen listesini temizlemeyi bırakmayı geçersiz kılar.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.gb_yeniGörev = New System.Windows.Forms.GroupBox()
        Me.cbl_personeller = New System.Windows.Forms.CheckedListBox()
        Me.cb_gorevYeri = New System.Windows.Forms.ComboBox()
        Me.lbl_gorevYeri = New System.Windows.Forms.Label()
        Me.cb_birim = New System.Windows.Forms.ComboBox()
        Me.lbl_birim = New System.Windows.Forms.Label()
        Me.dtp_baslangic = New System.Windows.Forms.DateTimePicker()
        Me.lbl_baslangic = New System.Windows.Forms.Label()
        Me.dtp_bitis = New System.Windows.Forms.DateTimePicker()
        Me.lbl_bitis = New System.Windows.Forms.Label()
        Me.lbl_personelSec = New System.Windows.Forms.Label()
        Me.cb_gorevTuru = New System.Windows.Forms.ComboBox()
        Me.lbl_gorevTuru = New System.Windows.Forms.Label()
        Me.tb_aciklama = New System.Windows.Forms.TextBox()
        Me.lbl_aciklama = New System.Windows.Forms.Label()
        Me.tb_gorevAdi = New System.Windows.Forms.TextBox()
        Me.lbl_gorevAdi = New System.Windows.Forms.Label()
        Me.btn_gorevEkle = New System.Windows.Forms.Button()
        Me.btn_personelBilgisi = New System.Windows.Forms.Button()
        Me.btn_bitenGorevler = New System.Windows.Forms.Button()
        Me.gb_gorevler = New System.Windows.Forms.GroupBox()
        Me.lbl_gorevAyrinti = New System.Windows.Forms.Label()
        Me.lbl_ayrinti = New System.Windows.Forms.Label()
        Me.lb_gorevler = New System.Windows.Forms.ListBox()
        Me.btn_cikis = New System.Windows.Forms.Button()
        Me.btn_yenile = New System.Windows.Forms.Button()
        Me.gb_yeniGörev.SuspendLayout()
        Me.gb_gorevler.SuspendLayout()
        Me.SuspendLayout()
        '
        'gb_yeniGörev
        '
        Me.gb_yeniGörev.Controls.Add(Me.cbl_personeller)
        Me.gb_yeniGörev.Controls.Add(Me.cb_gorevYeri)
        Me.gb_yeniGörev.Controls.Add(Me.lbl_gorevYeri)
        Me.gb_yeniGörev.Controls.Add(Me.cb_birim)
        Me.gb_yeniGörev.Controls.Add(Me.lbl_birim)
        Me.gb_yeniGörev.Controls.Add(Me.dtp_baslangic)
        Me.gb_yeniGörev.Controls.Add(Me.lbl_baslangic)
        Me.gb_yeniGörev.Controls.Add(Me.dtp_bitis)
        Me.gb_yeniGörev.Controls.Add(Me.lbl_bitis)
        Me.gb_yeniGörev.Controls.Add(Me.lbl_personelSec)
        Me.gb_yeniGörev.Controls.Add(Me.cb_gorevTuru)
        Me.gb_yeniGörev.Controls.Add(Me.lbl_gorevTuru)
        Me.gb_yeniGörev.Controls.Add(Me.tb_aciklama)
        Me.gb_yeniGörev.Controls.Add(Me.lbl_aciklama)
        Me.gb_yeniGörev.Controls.Add(Me.tb_gorevAdi)
        Me.gb_yeniGörev.Controls.Add(Me.lbl_gorevAdi)
        Me.gb_yeniGörev.Controls.Add(Me.btn_gorevEkle)
        Me.gb_yeniGörev.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.gb_yeniGörev.Location = New System.Drawing.Point(12, 74)
        Me.gb_yeniGörev.Name = "gb_yeniGörev"
        Me.gb_yeniGörev.Size = New System.Drawing.Size(335, 428)
        Me.gb_yeniGörev.TabIndex = 0
        Me.gb_yeniGörev.TabStop = False
        Me.gb_yeniGörev.Text = "Görev Oluştur"
        '
        'cbl_personeller
        '
        Me.cbl_personeller.FormattingEnabled = True
        Me.cbl_personeller.Location = New System.Drawing.Point(86, 251)
        Me.cbl_personeller.Name = "cbl_personeller"
        Me.cbl_personeller.Size = New System.Drawing.Size(199, 55)
        Me.cbl_personeller.TabIndex = 9
        '
        'cb_gorevYeri
        '
        Me.cb_gorevYeri.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cb_gorevYeri.FormattingEnabled = True
        Me.cb_gorevYeri.Location = New System.Drawing.Point(88, 212)
        Me.cb_gorevYeri.Name = "cb_gorevYeri"
        Me.cb_gorevYeri.Size = New System.Drawing.Size(121, 24)
        Me.cb_gorevYeri.TabIndex = 4
        '
        'lbl_gorevYeri
        '
        Me.lbl_gorevYeri.AutoSize = True
        Me.lbl_gorevYeri.Location = New System.Drawing.Point(6, 215)
        Me.lbl_gorevYeri.Name = "lbl_gorevYeri"
        Me.lbl_gorevYeri.Size = New System.Drawing.Size(78, 16)
        Me.lbl_gorevYeri.TabIndex = 0
        Me.lbl_gorevYeri.Text = "Görev Yeri :"
        '
        'cb_birim
        '
        Me.cb_birim.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cb_birim.FormattingEnabled = True
        Me.cb_birim.Location = New System.Drawing.Point(88, 182)
        Me.cb_birim.Name = "cb_birim"
        Me.cb_birim.Size = New System.Drawing.Size(121, 24)
        Me.cb_birim.TabIndex = 4
        '
        'lbl_birim
        '
        Me.lbl_birim.AutoSize = True
        Me.lbl_birim.Location = New System.Drawing.Point(6, 185)
        Me.lbl_birim.Name = "lbl_birim"
        Me.lbl_birim.Size = New System.Drawing.Size(44, 16)
        Me.lbl_birim.TabIndex = 0
        Me.lbl_birim.Text = "Birim :"
        '
        'dtp_baslangic
        '
        Me.dtp_baslangic.Location = New System.Drawing.Point(85, 320)
        Me.dtp_baslangic.Name = "dtp_baslangic"
        Me.dtp_baslangic.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dtp_baslangic.Size = New System.Drawing.Size(200, 22)
        Me.dtp_baslangic.TabIndex = 6
        '
        'lbl_baslangic
        '
        Me.lbl_baslangic.Location = New System.Drawing.Point(3, 320)
        Me.lbl_baslangic.Name = "lbl_baslangic"
        Me.lbl_baslangic.Size = New System.Drawing.Size(76, 36)
        Me.lbl_baslangic.TabIndex = 0
        Me.lbl_baslangic.Text = "Başlangıç Tarihi :"
        '
        'dtp_bitis
        '
        Me.dtp_bitis.Location = New System.Drawing.Point(85, 362)
        Me.dtp_bitis.Name = "dtp_bitis"
        Me.dtp_bitis.Size = New System.Drawing.Size(200, 22)
        Me.dtp_bitis.TabIndex = 7
        '
        'lbl_bitis
        '
        Me.lbl_bitis.AutoSize = True
        Me.lbl_bitis.Location = New System.Drawing.Point(3, 367)
        Me.lbl_bitis.Name = "lbl_bitis"
        Me.lbl_bitis.Size = New System.Drawing.Size(76, 16)
        Me.lbl_bitis.TabIndex = 0
        Me.lbl_bitis.Text = "Bitiş Tarihi :"
        '
        'lbl_personelSec
        '
        Me.lbl_personelSec.AutoSize = True
        Me.lbl_personelSec.Location = New System.Drawing.Point(3, 254)
        Me.lbl_personelSec.Name = "lbl_personelSec"
        Me.lbl_personelSec.Size = New System.Drawing.Size(68, 16)
        Me.lbl_personelSec.TabIndex = 0
        Me.lbl_personelSec.Text = "Personel :"
        '
        'cb_gorevTuru
        '
        Me.cb_gorevTuru.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cb_gorevTuru.FormattingEnabled = True
        Me.cb_gorevTuru.Location = New System.Drawing.Point(88, 152)
        Me.cb_gorevTuru.Name = "cb_gorevTuru"
        Me.cb_gorevTuru.Size = New System.Drawing.Size(121, 24)
        Me.cb_gorevTuru.TabIndex = 3
        '
        'lbl_gorevTuru
        '
        Me.lbl_gorevTuru.AutoSize = True
        Me.lbl_gorevTuru.Location = New System.Drawing.Point(6, 155)
        Me.lbl_gorevTuru.Name = "lbl_gorevTuru"
        Me.lbl_gorevTuru.Size = New System.Drawing.Size(81, 16)
        Me.lbl_gorevTuru.TabIndex = 0
        Me.lbl_gorevTuru.Text = "Görev Türü :"
        '
        'tb_aciklama
        '
        Me.tb_aciklama.Location = New System.Drawing.Point(88, 50)
        Me.tb_aciklama.Multiline = True
        Me.tb_aciklama.Name = "tb_aciklama"
        Me.tb_aciklama.Size = New System.Drawing.Size(239, 93)
        Me.tb_aciklama.TabIndex = 2
        '
        'lbl_aciklama
        '
        Me.lbl_aciklama.AutoSize = True
        Me.lbl_aciklama.Location = New System.Drawing.Point(8, 53)
        Me.lbl_aciklama.Name = "lbl_aciklama"
        Me.lbl_aciklama.Size = New System.Drawing.Size(70, 16)
        Me.lbl_aciklama.TabIndex = 0
        Me.lbl_aciklama.Text = "Açıklama :"
        '
        'tb_gorevAdi
        '
        Me.tb_gorevAdi.Location = New System.Drawing.Point(88, 22)
        Me.tb_gorevAdi.Name = "tb_gorevAdi"
        Me.tb_gorevAdi.Size = New System.Drawing.Size(239, 22)
        Me.tb_gorevAdi.TabIndex = 1
        '
        'lbl_gorevAdi
        '
        Me.lbl_gorevAdi.AutoSize = True
        Me.lbl_gorevAdi.Location = New System.Drawing.Point(8, 25)
        Me.lbl_gorevAdi.Name = "lbl_gorevAdi"
        Me.lbl_gorevAdi.Size = New System.Drawing.Size(74, 16)
        Me.lbl_gorevAdi.TabIndex = 0
        Me.lbl_gorevAdi.Text = "Görev Adı :"
        '
        'btn_gorevEkle
        '
        Me.btn_gorevEkle.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.btn_gorevEkle.Location = New System.Drawing.Point(198, 390)
        Me.btn_gorevEkle.Name = "btn_gorevEkle"
        Me.btn_gorevEkle.Size = New System.Drawing.Size(87, 32)
        Me.btn_gorevEkle.TabIndex = 8
        Me.btn_gorevEkle.Text = "Görev Ekle"
        Me.btn_gorevEkle.UseVisualStyleBackColor = True
        '
        'btn_personelBilgisi
        '
        Me.btn_personelBilgisi.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.btn_personelBilgisi.Location = New System.Drawing.Point(572, 12)
        Me.btn_personelBilgisi.Name = "btn_personelBilgisi"
        Me.btn_personelBilgisi.Size = New System.Drawing.Size(122, 32)
        Me.btn_personelBilgisi.TabIndex = 9
        Me.btn_personelBilgisi.Text = "Personel Bilgileri"
        Me.btn_personelBilgisi.UseVisualStyleBackColor = True
        '
        'btn_bitenGorevler
        '
        Me.btn_bitenGorevler.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.btn_bitenGorevler.Location = New System.Drawing.Point(700, 12)
        Me.btn_bitenGorevler.Name = "btn_bitenGorevler"
        Me.btn_bitenGorevler.Size = New System.Drawing.Size(122, 32)
        Me.btn_bitenGorevler.TabIndex = 10
        Me.btn_bitenGorevler.Text = "Biten Görevler"
        Me.btn_bitenGorevler.UseVisualStyleBackColor = True
        '
        'gb_gorevler
        '
        Me.gb_gorevler.Controls.Add(Me.lbl_gorevAyrinti)
        Me.gb_gorevler.Controls.Add(Me.lbl_ayrinti)
        Me.gb_gorevler.Controls.Add(Me.lb_gorevler)
        Me.gb_gorevler.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.gb_gorevler.Location = New System.Drawing.Point(353, 74)
        Me.gb_gorevler.Name = "gb_gorevler"
        Me.gb_gorevler.Size = New System.Drawing.Size(469, 428)
        Me.gb_gorevler.TabIndex = 0
        Me.gb_gorevler.TabStop = False
        Me.gb_gorevler.Text = "Görevler"
        '
        'lbl_gorevAyrinti
        '
        Me.lbl_gorevAyrinti.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbl_gorevAyrinti.Location = New System.Drawing.Point(194, 38)
        Me.lbl_gorevAyrinti.Name = "lbl_gorevAyrinti"
        Me.lbl_gorevAyrinti.Size = New System.Drawing.Size(269, 372)
        Me.lbl_gorevAyrinti.TabIndex = 0
        '
        'lbl_ayrinti
        '
        Me.lbl_ayrinti.AutoSize = True
        Me.lbl_ayrinti.Location = New System.Drawing.Point(191, 0)
        Me.lbl_ayrinti.Name = "lbl_ayrinti"
        Me.lbl_ayrinti.Size = New System.Drawing.Size(102, 16)
        Me.lbl_ayrinti.TabIndex = 0
        Me.lbl_ayrinti.Text = "Görev Ayrıntıları"
        '
        'lb_gorevler
        '
        Me.lb_gorevler.FormattingEnabled = True
        Me.lb_gorevler.ItemHeight = 16
        Me.lb_gorevler.Location = New System.Drawing.Point(7, 38)
        Me.lb_gorevler.Name = "lb_gorevler"
        Me.lb_gorevler.Size = New System.Drawing.Size(181, 372)
        Me.lb_gorevler.TabIndex = 0
        '
        'btn_cikis
        '
        Me.btn_cikis.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.btn_cikis.Location = New System.Drawing.Point(12, 12)
        Me.btn_cikis.Name = "btn_cikis"
        Me.btn_cikis.Size = New System.Drawing.Size(122, 32)
        Me.btn_cikis.TabIndex = 10
        Me.btn_cikis.Text = "Çıkış Yap"
        Me.btn_cikis.UseVisualStyleBackColor = True
        '
        'btn_yenile
        '
        Me.btn_yenile.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.btn_yenile.Location = New System.Drawing.Point(444, 12)
        Me.btn_yenile.Name = "btn_yenile"
        Me.btn_yenile.Size = New System.Drawing.Size(122, 32)
        Me.btn_yenile.TabIndex = 9
        Me.btn_yenile.Text = "Yenile"
        Me.btn_yenile.UseVisualStyleBackColor = True
        '
        'yoneticiAnasayfa
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(834, 514)
        Me.Controls.Add(Me.gb_gorevler)
        Me.Controls.Add(Me.btn_cikis)
        Me.Controls.Add(Me.btn_bitenGorevler)
        Me.Controls.Add(Me.btn_yenile)
        Me.Controls.Add(Me.btn_personelBilgisi)
        Me.Controls.Add(Me.gb_yeniGörev)
        Me.MaximizeBox = False
        Me.Name = "yoneticiAnasayfa"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Yönetici Anasayfa"
        Me.gb_yeniGörev.ResumeLayout(False)
        Me.gb_yeniGörev.PerformLayout()
        Me.gb_gorevler.ResumeLayout(False)
        Me.gb_gorevler.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents gb_yeniGörev As GroupBox
    Friend WithEvents btn_gorevEkle As Button
    Friend WithEvents cb_gorevTuru As ComboBox
    Friend WithEvents lbl_gorevTuru As Label
    Friend WithEvents tb_aciklama As TextBox
    Friend WithEvents lbl_aciklama As Label
    Friend WithEvents tb_gorevAdi As TextBox
    Friend WithEvents lbl_gorevAdi As Label
    Friend WithEvents dtp_bitis As DateTimePicker
    Friend WithEvents lbl_bitis As Label
    Friend WithEvents lbl_personelSec As Label
    Friend WithEvents dtp_baslangic As DateTimePicker
    Friend WithEvents lbl_baslangic As Label
    Friend WithEvents cb_birim As ComboBox
    Friend WithEvents lbl_birim As Label
    Friend WithEvents btn_personelBilgisi As Button
    Friend WithEvents btn_bitenGorevler As Button
    Friend WithEvents gb_gorevler As GroupBox
    Friend WithEvents lb_gorevler As ListBox
    Friend WithEvents lbl_ayrinti As Label
    Friend WithEvents lbl_gorevAyrinti As Label
    Friend WithEvents btn_cikis As Button
    Friend WithEvents cbl_personeller As CheckedListBox
    Friend WithEvents cb_gorevYeri As ComboBox
    Friend WithEvents lbl_gorevYeri As Label
    Friend WithEvents btn_yenile As Button
End Class
