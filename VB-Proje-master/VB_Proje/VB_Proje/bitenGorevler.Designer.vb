<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class bitenGorevler
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
        Me.lbl_bitenGorevler = New System.Windows.Forms.Label()
        Me.lb_bitenGorevler = New System.Windows.Forms.ListBox()
        Me.btn_rapor = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lbl_bitenGorevler
        '
        Me.lbl_bitenGorevler.AutoSize = True
        Me.lbl_bitenGorevler.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.lbl_bitenGorevler.Location = New System.Drawing.Point(31, 29)
        Me.lbl_bitenGorevler.Name = "lbl_bitenGorevler"
        Me.lbl_bitenGorevler.Size = New System.Drawing.Size(93, 16)
        Me.lbl_bitenGorevler.TabIndex = 0
        Me.lbl_bitenGorevler.Text = "Biten Görevler"
        '
        'lb_bitenGorevler
        '
        Me.lb_bitenGorevler.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.lb_bitenGorevler.FormattingEnabled = True
        Me.lb_bitenGorevler.ItemHeight = 16
        Me.lb_bitenGorevler.Items.AddRange(New Object() {"Personel ->Biten Görev 1", "Personel ->Biten Görev 2", "Personel ->Biten Görev 3", "Personel ->Biten Görev 4"})
        Me.lb_bitenGorevler.Location = New System.Drawing.Point(34, 57)
        Me.lb_bitenGorevler.Name = "lb_bitenGorevler"
        Me.lb_bitenGorevler.Size = New System.Drawing.Size(267, 356)
        Me.lb_bitenGorevler.TabIndex = 1
        '
        'btn_rapor
        '
        Me.btn_rapor.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.btn_rapor.Location = New System.Drawing.Point(203, 419)
        Me.btn_rapor.Name = "btn_rapor"
        Me.btn_rapor.Size = New System.Drawing.Size(98, 31)
        Me.btn_rapor.TabIndex = 2
        Me.btn_rapor.Text = "Rapor Oluştur"
        Me.btn_rapor.UseVisualStyleBackColor = True
        '
        'bitenGorevler
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(342, 459)
        Me.Controls.Add(Me.btn_rapor)
        Me.Controls.Add(Me.lb_bitenGorevler)
        Me.Controls.Add(Me.lbl_bitenGorevler)
        Me.MaximizeBox = False
        Me.Name = "bitenGorevler"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Biten Görevler"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lbl_bitenGorevler As Label
    Friend WithEvents lb_bitenGorevler As ListBox
    Friend WithEvents btn_rapor As Button
End Class
