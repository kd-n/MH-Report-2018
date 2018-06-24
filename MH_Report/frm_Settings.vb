Imports System.IO

Public Class frm_Settings

    Dim flgTeam As Boolean
    Dim flgSource As Boolean
    Dim flgOutput As Boolean
    Dim lastUpdate As Date
    Dim directoryPath As String

    Private Sub frm_Settings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' retrieve user settings and set to textbox
        txt_SourceFolder.Text = My.Settings.SourceFolder
        txt_OutputFolder.Text = My.Settings.OutputFolder
        txt_TeamListUpdate.Text = "Last updated on " + My.Settings.TeamListUpdate
    End Sub

    Private Sub btn_SourceChange_Click(sender As Object, e As EventArgs) Handles btn_SourceChange.Click
        Dim fbd As FolderBrowserDialog = New FolderBrowserDialog()

        ' Set the Help text description for the FolderBrowserDialog. 
        fbd.Description = "Please select Source folder"
        ' display dialog box
        fbd.ShowDialog()

        ' check changes in the selected path
        If Not fbd.SelectedPath() = "" AndAlso Not txt_SourceFolder.Text = fbd.SelectedPath() Then
            ' set selected path to textbox
            txt_SourceFolder.Text = fbd.SelectedPath()
            flgSource = True
        End If

        fbd = Nothing
    End Sub

    Private Sub btn_Output_Click(sender As Object, e As EventArgs) Handles btn_Output.Click
        Dim fbd As FolderBrowserDialog = New FolderBrowserDialog()

        ' Set the Help text description for the FolderBrowserDialog. 
        fbd.Description = "Please select Output folder"
        ' display dialog box
        fbd.ShowDialog()

        ' check changes in the selected path
        If Not fbd.SelectedPath() = "" AndAlso Not txt_OutputFolder.Text = fbd.SelectedPath() Then
            ' set selected path to textbox
            txt_OutputFolder.Text = fbd.SelectedPath()
            flgOutput = True
        End If

        fbd = Nothing
    End Sub

    Private Sub btn_Team_Click(sender As Object, e As EventArgs) Handles btn_Team.Click
        Dim openFile As OpenFileDialog = New OpenFileDialog

        ' Set filter options and filter index.
        openFile.Filter = "Excel Files (*.xls, *.xlsx)|*.xls;*.xlsx"
        openFile.FilterIndex = 1

        ' Call ShowDialog.
        Dim result As DialogResult = openFile.ShowDialog()

        ' Test result.
        If result = Windows.Forms.DialogResult.OK Then
            ' check the file type
            If openFile.FileName.Contains(".xls") OrElse directoryPath.Contains(".xlsx") Then
                ' save the directory path if an excel file
                directoryPath = openFile.FileName
                txt_TeamListUpdate.Text = directoryPath
                My.Settings.MyExcel = txt_TeamListUpdate.Text
            End If

            If Not openFile.FileName.Contains(".xls") AndAlso Not directoryPath.Contains(".xlsx") Then
                ' display error message if not an excel file
                MessageBox.Show("Incorrect file type.  Please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                ' save the directory path if an excel file
                directoryPath = openFile.FileName
                txt_TeamListUpdate.Text = directoryPath
                My.Settings.MyExcel = txt_TeamListUpdate.Text
            End If
        End If
    End Sub

    Private Sub btn_Save_Click(sender As Object, e As EventArgs) Handles btn_Save.Click
        Dim output As Integer

        ' set settings to application properties
        My.Settings.SourceFolder = txt_SourceFolder.Text
        My.Settings.OutputFolder = txt_OutputFolder.Text

        ' if path is not empty, create the resource file
        If Not directoryPath = Nothing Then
            ' set the cursor to wait while processing
            Me.Cursor = Cursors.WaitCursor
            ' disable all buttong while processing
            btn_SourceChange.Enabled = False
            btn_Output.Enabled = False
            btn_Team.Enabled = False
            btn_Save.Enabled = False
            btn_Close.Enabled = False
            output = TeamResource.UpdateTeamList(directoryPath)
        End If
        'set cursor back to default
        Me.Cursor = Cursors.Default

        ' enable all buttons
        btn_SourceChange.Enabled = True
        btn_Output.Enabled = True
        btn_Team.Enabled = True
        btn_Save.Enabled = True
        btn_Close.Enabled = True

        ' result is 0, processing is successful
        If output.Equals(0) Then
            flgTeam = True
            ' set date and time to textbox
            My.Settings.TeamListUpdate = DateTime.Now
        End If
        ' result is -1, display error message
        If output.Equals(-1) Then
            MessageBox.Show("An error occured during processing.  Please check the file and try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' check if all application settings are set
        If Not My.Settings.SourceFolder = Nothing AndAlso Not My.Settings.OutputFolder = Nothing AndAlso Not My.Settings.TeamListUpdate = Nothing AndAlso Not My.Settings.MyExcel = Nothing Then
            ' enable button
            frm_MainMenu.btn_MHData.Enabled = True
            ' display error message
            MessageBox.Show("Settings saved successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
        ' check if any of the application settings were not set
        If My.Settings.SourceFolder = Nothing OrElse My.Settings.OutputFolder = Nothing OrElse My.Settings.TeamListUpdate = Nothing OrElse My.Settings.MyExcel = Nothing Then
            ' disable button
            frm_MainMenu.btn_MHData.Enabled = False
            ' display error message
            MessageBox.Show("Settings were not configured.  Please click Change Settings to configure.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

        ' save settings
        My.Settings.Save()

        ' close form
        Me.Close()
    End Sub

    Private Sub btn_Close_Click(sender As Object, e As EventArgs) Handles btn_Close.Click
        Dim dialog As DialogResult = New DialogResult

        If flgSource OrElse flgOutput OrElse Not directoryPath = Nothing Then
            dialog = MessageBox.Show("Changes have been made.  Exit without saving changes?", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If dialog = Windows.Forms.DialogResult.No Then
                Return
            End If
        End If
        Me.Close()
    End Sub

End Class