<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class kullaniciGiris_ekrani
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
        Me.btn_giris = New System.Windows.Forms.Button()
        Me.tb_kullaniciAdi = New System.Windows.Forms.TextBox()
        Me.tb_sifre = New System.Windows.Forms.TextBox()
        Me.lbl_kullaniciAdi = New System.Windows.Forms.Label()
        Me.lbl_sifre = New System.Windows.Forms.Label()
        Me.cb_hatirla = New System.Windows.Forms.CheckBox()
        Me.gb_giris = New System.Windows.Forms.GroupBox()
        Me.gb_giris.SuspendLayout()
        Me.SuspendLayout()
        '
        'btn_giris
        '
        Me.btn_giris.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.btn_giris.Location = New System.Drawing.Point(156, 154)
        Me.btn_giris.Name = "btn_giris"
        Me.btn_giris.Size = New System.Drawing.Size(89, 35)
        Me.btn_giris.TabIndex = 3
        Me.btn_giris.Text = "Giriş"
        Me.btn_giris.UseVisualStyleBackColor = True
        '
        'tb_kullaniciAdi
        '
        Me.tb_kullaniciAdi.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.tb_kullaniciAdi.Location = New System.Drawing.Point(51, 58)
        Me.tb_kullaniciAdi.MaxLength = 50
        Me.tb_kullaniciAdi.Name = "tb_kullaniciAdi"
        Me.tb_kullaniciAdi.Size = New System.Drawing.Size(194, 22)
        Me.tb_kullaniciAdi.TabIndex = 1
        '
        'tb_sifre
        '
        Me.tb_sifre.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.tb_sifre.Location = New System.Drawing.Point(51, 126)
        Me.tb_sifre.MaxLength = 50
        Me.tb_sifre.Name = "tb_sifre"
        Me.tb_sifre.Size = New System.Drawing.Size(194, 22)
        Me.tb_sifre.TabIndex = 2
        Me.tb_sifre.UseSystemPasswordChar = True
        '
        'lbl_kullaniciAdi
        '
        Me.lbl_kullaniciAdi.AutoSize = True
        Me.lbl_kullaniciAdi.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.lbl_kullaniciAdi.Location = New System.Drawing.Point(48, 39)
        Me.lbl_kullaniciAdi.Name = "lbl_kullaniciAdi"
        Me.lbl_kullaniciAdi.Size = New System.Drawing.Size(80, 16)
        Me.lbl_kullaniciAdi.TabIndex = 2
        Me.lbl_kullaniciAdi.Text = "Kullanıcı Adı"
        '
        'lbl_sifre
        '
        Me.lbl_sifre.AutoSize = True
        Me.lbl_sifre.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.lbl_sifre.Location = New System.Drawing.Point(48, 98)
        Me.lbl_sifre.Name = "lbl_sifre"
        Me.lbl_sifre.Size = New System.Drawing.Size(35, 16)
        Me.lbl_sifre.TabIndex = 2
        Me.lbl_sifre.Text = "Şifre"
        '
        'cb_hatirla
        '
        Me.cb_hatirla.AutoSize = True
        Me.cb_hatirla.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.cb_hatirla.Location = New System.Drawing.Point(51, 154)
        Me.cb_hatirla.Name = "cb_hatirla"
        Me.cb_hatirla.Size = New System.Drawing.Size(96, 20)
        Me.cb_hatirla.TabIndex = 4
        Me.cb_hatirla.Text = "Beni Hatırla"
        Me.cb_hatirla.UseVisualStyleBackColor = True
        '
        'gb_giris
        '
        Me.gb_giris.Controls.Add(Me.tb_kullaniciAdi)
        Me.gb_giris.Controls.Add(Me.tb_sifre)
        Me.gb_giris.Controls.Add(Me.cb_hatirla)
        Me.gb_giris.Controls.Add(Me.lbl_sifre)
        Me.gb_giris.Controls.Add(Me.btn_giris)
        Me.gb_giris.Controls.Add(Me.lbl_kullaniciAdi)
        Me.gb_giris.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.gb_giris.Location = New System.Drawing.Point(46, 43)
        Me.gb_giris.Name = "gb_giris"
        Me.gb_giris.Size = New System.Drawing.Size(290, 225)
        Me.gb_giris.TabIndex = 5
        Me.gb_giris.TabStop = False
        Me.gb_giris.Text = "Kullanıcı Girişi"
        '
        'kullaniciGiris_ekrani
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(384, 311)
        Me.Controls.Add(Me.gb_giris)
        Me.MaximizeBox = False
        Me.Name = "kullaniciGiris_ekrani"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.gb_giris.ResumeLayout(False)
        Me.gb_giris.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btn_giris As Button
    Friend WithEvents tb_kullaniciAdi As TextBox
    Friend WithEvents tb_sifre As TextBox
    Friend WithEvents lbl_kullaniciAdi As Label
    Friend WithEvents lbl_sifre As Label
    Friend WithEvents cb_hatirla As CheckBox
    Friend WithEvents gb_giris As GroupBox
End Class
