<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_graphtry
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
        Me.btn_UploadEraj = New System.Windows.Forms.Button()
        Me.txtbox_Eraj = New System.Windows.Forms.TextBox()
        Me.txt_totSource = New System.Windows.Forms.TextBox()
        Me.btn_uploadSource = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btn_UploadEraj
        '
        Me.btn_UploadEraj.Location = New System.Drawing.Point(307, 88)
        Me.btn_UploadEraj.Name = "btn_UploadEraj"
        Me.btn_UploadEraj.Size = New System.Drawing.Size(75, 23)
        Me.btn_UploadEraj.TabIndex = 10
        Me.btn_UploadEraj.Text = "Upload"
        Me.btn_UploadEraj.UseVisualStyleBackColor = True
        '
        'txtbox_Eraj
        '
        Me.txtbox_Eraj.Location = New System.Drawing.Point(12, 90)
        Me.txtbox_Eraj.Name = "txtbox_Eraj"
        Me.txtbox_Eraj.ReadOnly = True
        Me.txtbox_Eraj.Size = New System.Drawing.Size(287, 20)
        Me.txtbox_Eraj.TabIndex = 11
        Me.txtbox_Eraj.TabStop = False
        '
        'txt_totSource
        '
        Me.txt_totSource.Location = New System.Drawing.Point(12, 40)
        Me.txt_totSource.Name = "txt_totSource"
        Me.txt_totSource.ReadOnly = True
        Me.txt_totSource.Size = New System.Drawing.Size(287, 20)
        Me.txt_totSource.TabIndex = 12
        Me.txt_totSource.TabStop = False
        '
        'btn_uploadSource
        '
        Me.btn_uploadSource.Location = New System.Drawing.Point(307, 40)
        Me.btn_uploadSource.Name = "btn_uploadSource"
        Me.btn_uploadSource.Size = New System.Drawing.Size(75, 23)
        Me.btn_uploadSource.TabIndex = 13
        Me.btn_uploadSource.Text = "Upload"
        Me.btn_uploadSource.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(106, 13)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Team Totals Source:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 74)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(121, 13)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "ERAJ Template Source:"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(71, 126)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(112, 23)
        Me.Button1.TabIndex = 16
        Me.Button1.Text = "Generate ERAJ"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(187, 126)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(112, 23)
        Me.Button2.TabIndex = 17
        Me.Button2.Text = "Back"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'frm_graphtry
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(394, 171)
        Me.ControlBox = False
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btn_uploadSource)
        Me.Controls.Add(Me.txt_totSource)
        Me.Controls.Add(Me.btn_UploadEraj)
        Me.Controls.Add(Me.txtbox_Eraj)
        Me.Name = "frm_graphtry"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Draw ERAJ"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btn_UploadEraj As System.Windows.Forms.Button
    Friend WithEvents txtbox_Eraj As System.Windows.Forms.TextBox
    Friend WithEvents txt_totSource As System.Windows.Forms.TextBox
    Friend WithEvents btn_uploadSource As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Button2 As System.Windows.Forms.Button
End Class
