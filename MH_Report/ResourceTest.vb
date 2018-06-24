Imports System.Resources

Module ResourceTest
    Private Const TEAM_RES As String = ".\teams.resx"
    Private Const TITLE As String = "Title"
    Private Const DESCRIPTION As String = "Description"
    Private Const TEAMS As String = "Teams"
    Private Const MEMBERS As String = "Members"
    Private Const TITLE_TXT As String = "Team Resource"
    Private Const DESC_TXT As String = "List of FETP Teams, Departments, Divisions, and their members"

    Public Sub GenerateTeamResource()
        Dim teamNames As List(Of String) ' list for team names
        Dim teamMembers As List(Of Team) = New List(Of Team) ' list for team members
        Dim team As Team

        ' set test values
        teamNames = New List(Of String) From {"SDT1", "SDT2", "SDT3"}

        team = New Team("SDT1", New List(Of String) From {"espie", "tholitz", "raymark"})
        teamMembers.Add(team)

        team = New Team("SDT2", New List(Of String) From {"tots", "norman", "ynna"})
        teamMembers.Add(team)

        team = New Team("SDT3", New List(Of String) From {"aimee", "eulah"})
        teamMembers.Add(team)

        Using teamRes As New ResXResourceWriter(TEAM_RES)
            ' title
            teamRes.AddResource(TITLE, TITLE_TXT)
            ' description
            teamRes.AddResource(DESCRIPTION, DESC_TXT)
            ' list of teams
            teamRes.AddResource(TEAMS, teamNames)
            ' list of members, grouped into teams
            teamRes.AddResource(MEMBERS, teamMembers)
            teamRes.Close()
        End Using
    End Sub

    Public Function TeamList() As List(Of String)
        Dim teamNames As List(Of String)

        Using teamResSet As New ResXResourceSet(TEAM_RES)
            teamNames = CType(teamResSet.GetObject(TEAMS), List(Of String))
            teamResSet.Close()
        End Using

        TeamList = teamNames
    End Function
End Module
