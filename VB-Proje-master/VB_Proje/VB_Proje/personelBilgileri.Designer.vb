<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class personelBilgileri
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
        Me.cb_personel = New System.Windows.Forms.ComboBox()
        Me.lbl_personel = New System.Windows.Forms.Label()
        Me.gb_gorevBilgileri = New System.Windows.Forms.GroupBox()
        Me.lbl_aktifGorev = New System.Windows.Forms.Label()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.lbl_bitenGorevSayisi = New System.Windows.Forms.Label()
        Me.lbl_bitenGorev = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.pb_doluluk = New System.Windows.Forms.PictureBox()
        Me.lbl_doluluk = New System.Windows.Forms.Label()
        Me.gb_gorevBilgileri.SuspendLayout()
        CType(Me.pb_doluluk, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cb_personel
        '
        Me.cb_personel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cb_personel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.cb_personel.FormattingEnabled = True
        Me.cb_personel.Items.AddRange(New Object() {"Personel 1", "Personel 2", "Personel 3", "Personel 4"})
        Me.cb_personel.Location = New System.Drawing.Point(88, 20)
        Me.cb_personel.Name = "cb_personel"
        Me.cb_personel.Size = New System.Drawing.Size(121, 24)
        Me.cb_personel.TabIndex = 1
        '
        'lbl_personel
        '
        Me.lbl_personel.AutoSize = True
        Me.lbl_personel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.lbl_personel.Location = New System.Drawing.Point(15, 23)
        Me.lbl_personel.Name = "lbl_personel"
        Me.lbl_personel.Size = New System.Drawing.Size(68, 16)
        Me.lbl_personel.TabIndex = 1
        Me.lbl_personel.Text = "Personel :"
        '
        'gb_gorevBilgileri
        '
        Me.gb_gorevBilgileri.Controls.Add(Me.lbl_aktifGorev)
        Me.gb_gorevBilgileri.Controls.Add(Me.ListBox1)
        Me.gb_gorevBilgileri.Controls.Add(Me.lbl_bitenGorevSayisi)
        Me.gb_gorevBilgileri.Controls.Add(Me.lbl_bitenGorev)
        Me.gb_gorevBilgileri.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.gb_gorevBilgileri.Location = New System.Drawing.Point(18, 129)
        Me.gb_gorevBilgileri.Name = "gb_gorevBilgileri"
        Me.gb_gorevBilgileri.Size = New System.Drawing.Size(339, 303)
        Me.gb_gorevBilgileri.TabIndex = 0
        Me.gb_gorevBilgileri.TabStop = False
        Me.gb_gorevBilgileri.Text = "Görev Bilgileri"
        '
        'lbl_aktifGorev
        '
        Me.lbl_aktifGorev.AutoSize = True
        Me.lbl_aktifGorev.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.lbl_aktifGorev.Location = New System.Drawing.Point(6, 63)
        Me.lbl_aktifGorev.Name = "lbl_aktifGorev"
        Me.lbl_aktifGorev.Size = New System.Drawing.Size(88, 16)
        Me.lbl_aktifGorev.TabIndex = 0
        Me.lbl_aktifGorev.Text = "Aktif Görevler"
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 16
        Me.ListBox1.Items.AddRange(New Object() {"Aktif Görev Adı - Bitiş Tarihi", "Aktif Görev Adı - Bitiş Tarihi", "Aktif Görev Adı - Bitiş Tarihi", "Aktif Görev Adı - Bitiş Tarihi"})
        Me.ListBox1.Location = New System.Drawing.Point(8, 82)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(325, 212)
        Me.ListBox1.TabIndex = 0
        '
        'lbl_bitenGorevSayisi
        '
        Me.lbl_bitenGorevSayisi.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.lbl_bitenGorevSayisi.Location = New System.Drawing.Point(148, 29)
        Me.lbl_bitenGorevSayisi.Name = "lbl_bitenGorevSayisi"
        Me.lbl_bitenGorevSayisi.Size = New System.Drawing.Size(43, 16)
        Me.lbl_bitenGorevSayisi.TabIndex = 0
        Me.lbl_bitenGorevSayisi.Text = "100"
        '
        'lbl_bitenGorev
        '
        Me.lbl_bitenGorev.AutoSize = True
        Me.lbl_bitenGorev.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.lbl_bitenGorev.Location = New System.Drawing.Point(5, 29)
        Me.lbl_bitenGorev.Name = "lbl_bitenGorev"
        Me.lbl_bitenGorev.Size = New System.Drawing.Size(137, 16)
        Me.lbl_bitenGorev.TabIndex = 0
        Me.lbl_bitenGorev.Text = "Bitirilen Görev Sayısı :"
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.Button1.Location = New System.Drawing.Point(249, 438)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(108, 32)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Rapor Oluştur"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'pb_doluluk
        '
        Me.pb_doluluk.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pb_doluluk.Image = Global.VB_Proje.My.Resources.Resources.yesilRenk
        Me.pb_doluluk.Location = New System.Drawing.Point(18, 66)
        Me.pb_doluluk.Name = "pb_doluluk"
        Me.pb_doluluk.Size = New System.Drawing.Size(340, 24)
        Me.pb_doluluk.TabIndex = 3
        Me.pb_doluluk.TabStop = False
        '
        'lbl_doluluk
        '
        Me.lbl_doluluk.AutoSize = True
        Me.lbl_doluluk.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.lbl_doluluk.Location = New System.Drawing.Point(108, 93)
        Me.lbl_doluluk.Name = "lbl_doluluk"
        Me.lbl_doluluk.Size = New System.Drawing.Size(165, 16)
        Me.lbl_doluluk.TabIndex = 4
        Me.lbl_doluluk.Text = "Personel Yoğunluğu %100"
        '
        'personelBilgileri
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(369, 491)
        Me.Controls.Add(Me.lbl_doluluk)
        Me.Controls.Add(Me.pb_doluluk)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.gb_gorevBilgileri)
        Me.Controls.Add(Me.lbl_personel)
        Me.Controls.Add(Me.cb_personel)
        Me.MaximizeBox = False
        Me.Name = "personelBilgileri"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Personel Bilgileri"
        Me.gb_gorevBilgileri.ResumeLayout(False)
        Me.gb_gorevBilgileri.PerformLayout()
        CType(Me.pb_doluluk, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cb_personel As ComboBox
    Friend WithEvents lbl_personel As Label
    Friend WithEvents gb_gorevBilgileri As GroupBox
    Friend WithEvents lbl_bitenGorevSayisi As Label
    Friend WithEvents lbl_bitenGorev As Label
    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents Button1 As Button
    Friend WithEvents lbl_aktifGorev As Label
    Friend WithEvents pb_doluluk As PictureBox
    Friend WithEvents lbl_doluluk As Label
End Class
