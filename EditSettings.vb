Imports System

Public Class Settings
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        My.Settings.WorkstationName = Me.WorkStationID.Text
        My.Settings.StartTime = Me.StartTime.Text
        My.Settings.EndTime = Me.EndTime.Text
        My.Settings.Freq = Me.Freq.Value
        My.Settings.Offset = Me.Offset.Value
        My.Settings.AutoStartOnLoad = Me.AutoStart.Checked
        My.Settings.AutoShutDown = Me.AutoShutDown.Checked
        My.Settings.FromUserName = Me.FromUserName.Text
        My.Settings.DistributeFromServer = Me.DistributeOnServer.Checked
        My.Settings.SMTPServerName = Me.SMTPServer.Text
        My.Settings.SMTPUser = Me.SMTPUserID.Text
        My.Settings.SMTPPassword = Me.SMTPPassword.Text
        Me.Close()
    End Sub

    Private Sub Settings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.WorkStationID.Text = My.Settings.WorkstationName
        Me.StartTime.Text = My.Settings.StartTime
        Me.EndTime.Text = My.Settings.EndTime
        Me.Offset.Value = CInt(My.Settings.Offset)
        Me.Freq.Value = CInt(My.Settings.Freq)
        Me.AutoStart.Checked = CBool(My.Settings.AutoStartOnLoad)
        Me.AutoShutDown.Checked = CBool(My.Settings.AutoShutDown)
        Me.FromUserName.Text = My.Settings.FromUserName
        Me.DistributeOnServer.Checked = CBool(My.Settings.DistributeFromServer)
        Me.SMTPServer.Text = My.Settings.SMTPServerName
        Me.SMTPUserID.Text = My.Settings.SMTPUser
        Me.SMTPPassword.Text = My.Settings.SMTPPassword
    End Sub
End Class