<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
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

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Letter_I = New System.Windows.Forms.TextBox()
        Me.Letter_S = New System.Windows.Forms.TextBox()
        Me.Letter_E = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.DefaultExt = "xls;xlsx"
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        Me.OpenFileDialog1.Filter = "Файлы Excel|*.xls;*.xlsx"
        Me.OpenFileDialog1.RestoreDirectory = True
        Me.OpenFileDialog1.Title = "Выбер файла для анализа"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(12, 12)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(374, 23)
        Me.ProgressBar1.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label1.Location = New System.Drawing.Point(12, 38)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(374, 29)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Введите буквы столбцов, если программа не смогла сама определить (названия не ста" &
    "ндартные)"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(323, 70)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(63, 99)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Рассчёт"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Letter_I
        '
        Me.Letter_I.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Letter_I.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Letter_I.Location = New System.Drawing.Point(245, 70)
        Me.Letter_I.Name = "Letter_I"
        Me.Letter_I.Size = New System.Drawing.Size(72, 29)
        Me.Letter_I.TabIndex = 3
        '
        'Letter_S
        '
        Me.Letter_S.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Letter_S.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Letter_S.Location = New System.Drawing.Point(245, 105)
        Me.Letter_S.Name = "Letter_S"
        Me.Letter_S.Size = New System.Drawing.Size(72, 29)
        Me.Letter_S.TabIndex = 3
        '
        'Letter_E
        '
        Me.Letter_E.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Letter_E.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Letter_E.Location = New System.Drawing.Point(245, 140)
        Me.Letter_E.Name = "Letter_E"
        Me.Letter_E.Size = New System.Drawing.Size(72, 29)
        Me.Letter_E.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 77)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(218, 18)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Буква столбца ""Исполнитель"""
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label3.Location = New System.Drawing.Point(12, 112)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(231, 18)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Буква столбца ""начало работы"""
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label4.Location = New System.Drawing.Point(12, 147)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(222, 18)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Буква столбца ""конец работы"""
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(401, 179)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Letter_E)
        Me.Controls.Add(Me.Letter_S)
        Me.Controls.Add(Me.Letter_I)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Name = "Form1"
        Me.Text = "Подсчёт неактивного времени"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents Label1 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents Letter_I As TextBox
    Friend WithEvents Letter_S As TextBox
    Friend WithEvents Letter_E As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
End Class
