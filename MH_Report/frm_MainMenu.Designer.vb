<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_MainMenu
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_MainMenu))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btn_MHData = New System.Windows.Forms.Button()
        Me.btn_Settings = New System.Windows.Forms.Button()
        Me.btn_Exit = New System.Windows.Forms.Button()
        Me.btn_drawEraj = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(90, 36)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(119, 26)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Main Menu"
        '
        'btn_MHData
        '
        Me.btn_MHData.Location = New System.Drawing.Point(95, 81)
        Me.btn_MHData.Name = "btn_MHData"
        Me.btn_MHData.Size = New System.Drawing.Size(110, 40)
        Me.btn_MHData.TabIndex = 1
        Me.btn_MHData.Text = "MH Data Report"
        Me.btn_MHData.UseVisualStyleBackColor = True
        '
        'btn_Settings
        '
        Me.btn_Settings.Location = New System.Drawing.Point(95, 144)
        Me.btn_Settings.Name = "btn_Settings"
        Me.btn_Settings.Size = New System.Drawing.Size(110, 40)
        Me.btn_Settings.TabIndex = 2
        Me.btn_Settings.Text = "Change Settings"
        Me.btn_Settings.UseVisualStyleBackColor = True
        '
        'btn_Exit
        '
        Me.btn_Exit.Location = New System.Drawing.Point(95, 258)
        Me.btn_Exit.Name = "btn_Exit"
        Me.btn_Exit.Size = New System.Drawing.Size(110, 40)
        Me.btn_Exit.TabIndex = 3
        Me.btn_Exit.Text = "Exit"
        Me.btn_Exit.UseVisualStyleBackColor = True
        '
        'btn_drawEraj
        '
        Me.btn_drawEraj.Location = New System.Drawing.Point(95, 199)
        Me.btn_drawEraj.Name = "btn_drawEraj"
        Me.btn_drawEraj.Size = New System.Drawing.Size(110, 40)
        Me.btn_drawEraj.TabIndex = 4
        Me.btn_drawEraj.Text = "Draw Graph"
        Me.btn_drawEraj.UseVisualStyleBackColor = True
        '
        'frm_MainMenu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(305, 349)
        Me.ControlBox = False
        Me.Controls.Add(Me.btn_drawEraj)
        Me.Controls.Add(Me.btn_Exit)
        Me.Controls.Add(Me.btn_Settings)
        Me.Controls.Add(Me.btn_MHData)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frm_MainMenu"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "MH Report"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btn_MHData As System.Windows.Forms.Button
    Friend WithEvents btn_Settings As System.Windows.Forms.Button
    Friend WithEvents btn_Exit As System.Windows.Forms.Button
    Friend WithEvents btn_drawEraj As System.Windows.Forms.Button

End Class
