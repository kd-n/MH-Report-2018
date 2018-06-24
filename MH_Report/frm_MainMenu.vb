Public Class frm_MainMenu

    Private Sub btn_Settings_Click(sender As Object, e As EventArgs) Handles btn_Settings.Click
        frm_settingsMenu.Show()
    End Sub

    Private Sub btn_Exit_Click(sender As Object, e As EventArgs) Handles btn_Exit.Click
        Me.Close()
    End Sub

    Private Sub btn_MHData_Click(sender As Object, e As EventArgs) Handles btn_MHData.Click
        Dim frmMHReport As frm_MHReport
        Me.Hide()
        frmMHReport = New frm_MHReport()
        frmMHReport.ShowDialog()
        frmMHReport = Nothing
        Me.Show()
    End Sub

    ' check if any of the application settings were not set
    Private Sub frm_MainMenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If My.Settings.SourceFolder = Nothing OrElse My.Settings.OutputFolder = Nothing OrElse My.Settings.TeamListUpdate = Nothing Then
            ' disable button
            btn_MHData.Enabled = False
            ' display error message if not an excel file
            MessageBox.Show("Settings were not configured.  Please click Change Settings to configure.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub btn_drawEraj_Click(sender As Object, e As EventArgs) Handles btn_drawEraj.Click
        frm_graphtry.Show()
    End Sub
End Class
