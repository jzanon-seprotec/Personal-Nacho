<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ListBox2 = New System.Windows.Forms.ListBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Versionlbl = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.SteelBlue
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 21.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(776, 67)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Drop xml file to panel and select language"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ComboBox1
        '
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"Arabic Algeria                   ar-dz", "Arabic EAU                        ar-ae", "Arabic Egipt                      ar-eg", "Arabic Kuwait                   ar-kw", "Arabic Morocco               ar-ma", "Arabic Saudi Arabian       ar-sa", "Bulgarian                           bg-gb", "Chinese Simplified           zh-cn", "Chinese Taiwan                 zh-tw", "Czeck                                 cs-cz", "Estonian                            et-ee", "German                             de-de", "Greek                                 el-gr", "Indonesian                        id-id", "Italian                                 it-it", "Japanese                            ja-jp", "Lithuanian                         lt-lt", "Persian Iran                       fa-ir", "Polish                                 pl-pl", "Portuguese                        pt-pt", "Romanian" & Global.Microsoft.VisualBasic.ChrW(9) & "                         ro-ro", "Russian                              ru-ru", "Serbian Latin                     sr-rs", "Slovak                                sk-sk", "Spanish Argentina           es-ar", "Spanish Chile                   es-cl", "Spanish Colombia           es-co", "Spanish Costa Rica          es-cr", "Spanish Mexico               es-mx", "Spanish Peru                    es-pe", "Spanish Spain                  es-es", "Spanish Venezuela          es-ve", "Thai                                   th-th", "Turkish                              tr-tr", "Ukrainian                          uk-ua", "Vietnamese                      vi-vn"})
        Me.ComboBox1.Location = New System.Drawing.Point(12, 100)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(775, 23)
        Me.ComboBox1.TabIndex = 1
        '
        'ListBox1
        '
        Me.ListBox1.AllowDrop = True
        Me.ListBox1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 15
        Me.ListBox1.Location = New System.Drawing.Point(12, 148)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(775, 109)
        Me.ListBox1.TabIndex = 2
        '
        'Button1
        '
        Me.Button1.Enabled = False
        Me.Button1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Button1.Location = New System.Drawing.Point(297, 387)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(188, 37)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Convert"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(9, 84)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(101, 15)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Select Language:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(9, 131)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 15)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Drop file here:"
        '
        'ListBox2
        '
        Me.ListBox2.FormattingEnabled = True
        Me.ListBox2.Location = New System.Drawing.Point(12, 280)
        Me.ListBox2.Name = "ListBox2"
        Me.ListBox2.Size = New System.Drawing.Size(775, 95)
        Me.ListBox2.TabIndex = 6
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(9, 262)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(30, 15)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Log:"
        '
        'Versionlbl
        '
        Me.Versionlbl.AutoSize = True
        Me.Versionlbl.Location = New System.Drawing.Point(703, 411)
        Me.Versionlbl.Name = "Versionlbl"
        Me.Versionlbl.Size = New System.Drawing.Size(84, 13)
        Me.Versionlbl.TabIndex = 8
        Me.Versionlbl.Text = "Version:             " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 436)
        Me.Controls.Add(Me.Versionlbl)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.ListBox2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(816, 475)
        Me.MinimumSize = New System.Drawing.Size(816, 475)
        Me.Name = "Form1"
        Me.Text = "WIPO XML Sequences"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents Button1 As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents ListBox2 As ListBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Versionlbl As Label
End Class
