<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_Settings
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_Settings))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txt_SourceFolder = New System.Windows.Forms.TextBox()
        Me.btn_SourceChange = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txt_OutputFolder = New System.Windows.Forms.TextBox()
        Me.btn_Output = New System.Windows.Forms.Button()
        Me.btn_Save = New System.Windows.Forms.Button()
        Me.btn_Close = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txt_TeamListUpdate = New System.Windows.Forms.TextBox()
        Me.btn_Team = New System.Windows.Forms.Button()
        Me.FileOpen = New System.Windows.Forms.OpenFileDialog()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(76, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Source Folder:"
        '
        'txt_SourceFolder
        '
        Me.txt_SourceFolder.Location = New System.Drawing.Point(12, 25)
        Me.txt_SourceFolder.Name = "txt_SourceFolder"
        Me.txt_SourceFolder.ReadOnly = True
        Me.txt_SourceFolder.Size = New System.Drawing.Size(349, 20)
        Me.txt_SourceFolder.TabIndex = 1
        Me.txt_SourceFolder.TabStop = False
        '
        'btn_SourceChange
        '
        Me.btn_SourceChange.Location = New System.Drawing.Point(367, 23)
        Me.btn_SourceChange.Name = "btn_SourceChange"
        Me.btn_SourceChange.Size = New System.Drawing.Size(75, 23)
        Me.btn_SourceChange.TabIndex = 2
        Me.btn_SourceChange.Text = "Change..."
        Me.btn_SourceChange.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 62)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(74, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Output Folder:"
        '
        'txt_OutputFolder
        '
        Me.txt_OutputFolder.Location = New System.Drawing.Point(12, 78)
        Me.txt_OutputFolder.Name = "txt_OutputFolder"
        Me.txt_OutputFolder.ReadOnly = True
        Me.txt_OutputFolder.Size = New System.Drawing.Size(349, 20)
        Me.txt_OutputFolder.TabIndex = 4
        Me.txt_OutputFolder.TabStop = False
        '
        'btn_Output
        '
        Me.btn_Output.Location = New System.Drawing.Point(367, 76)
        Me.btn_Output.Name = "btn_Output"
        Me.btn_Output.Size = New System.Drawing.Size(75, 23)
        Me.btn_Output.TabIndex = 3
        Me.btn_Output.Text = "Change..."
        Me.btn_Output.UseVisualStyleBackColor = True
        '
        'btn_Save
        '
        Me.btn_Save.Location = New System.Drawing.Point(130, 189)
        Me.btn_Save.Name = "btn_Save"
        Me.btn_Save.Size = New System.Drawing.Size(90, 25)
        Me.btn_Save.TabIndex = 4
        Me.btn_Save.Text = "Save Changes"
        Me.btn_Save.UseVisualStyleBackColor = True
        '
        'btn_Close
        '
        Me.btn_Close.Location = New System.Drawing.Point(235, 189)
        Me.btn_Close.Name = "btn_Close"
        Me.btn_Close.Size = New System.Drawing.Size(90, 25)
        Me.btn_Close.TabIndex = 1
        Me.btn_Close.Text = "Back"
        Me.btn_Close.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 118)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(157, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Team/Department Member List:"
        '
        'txt_TeamListUpdate
        '
        Me.txt_TeamListUpdate.Location = New System.Drawing.Point(12, 134)
        Me.txt_TeamListUpdate.Name = "txt_TeamListUpdate"
        Me.txt_TeamListUpdate.ReadOnly = True
        Me.txt_TeamListUpdate.Size = New System.Drawing.Size(349, 20)
        Me.txt_TeamListUpdate.TabIndex = 6
        Me.txt_TeamListUpdate.TabStop = False
        '
        'btn_Team
        '
        Me.btn_Team.Location = New System.Drawing.Point(367, 131)
        Me.btn_Team.Name = "btn_Team"
        Me.btn_Team.Size = New System.Drawing.Size(75, 23)
        Me.btn_Team.TabIndex = 7
        Me.btn_Team.Text = "Update..."
        Me.btn_Team.UseVisualStyleBackColor = True
        '
        'FileOpen
        '
        Me.FileOpen.FileName = "OpenFileDialog1"
        '
        'frm_Settings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(454, 226)
        Me.ControlBox = False
        Me.Controls.Add(Me.btn_Team)
        Me.Controls.Add(Me.txt_TeamListUpdate)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btn_Close)
        Me.Controls.Add(Me.btn_Save)
        Me.Controls.Add(Me.btn_Output)
        Me.Controls.Add(Me.txt_OutputFolder)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btn_SourceChange)
        Me.Controls.Add(Me.txt_SourceFolder)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frm_Settings"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Change Settings"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txt_SourceFolder As System.Windows.Forms.TextBox
    Friend WithEvents btn_SourceChange As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txt_OutputFolder As System.Windows.Forms.TextBox
    Friend WithEvents btn_Output As System.Windows.Forms.Button
    Friend WithEvents btn_Save As System.Windows.Forms.Button
    Friend WithEvents btn_Close As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txt_TeamListUpdate As System.Windows.Forms.TextBox
    Friend WithEvents btn_Team As System.Windows.Forms.Button
    Friend WithEvents FileOpen As System.Windows.Forms.OpenFileDialog
End Class
