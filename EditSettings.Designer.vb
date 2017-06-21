<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Settings
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
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.WorkStationID = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Freq = New System.Windows.Forms.NumericUpDown()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.StartTime = New System.Windows.Forms.TextBox()
        Me.EndTime = New System.Windows.Forms.TextBox()
        Me.Offset = New System.Windows.Forms.NumericUpDown()
        Me.AutoStart = New System.Windows.Forms.CheckBox()
        Me.AutoShutDown = New System.Windows.Forms.CheckBox()
        Me.RunOnTimer = New System.Windows.Forms.CheckBox()
        Me.FromUserName = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.DistributeOnServer = New System.Windows.Forms.CheckBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.SMTPPassword = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.SMTPUserID = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.SMTPServer = New System.Windows.Forms.TextBox()
        CType(Me.Freq, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Offset, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(254, 411)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 26)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Save"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(348, 411)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 26)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Cancel"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'WorkStationID
        '
        Me.WorkStationID.Location = New System.Drawing.Point(362, 12)
        Me.WorkStationID.Name = "WorkStationID"
        Me.WorkStationID.Size = New System.Drawing.Size(97, 20)
        Me.WorkStationID.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(247, 15)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(78, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Workstation ID"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(154, 96)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(52, 13)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "StartTime"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(154, 128)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(52, 13)
        Me.Label5.TabIndex = 5
        Me.Label5.Text = "End Time"
        '
        'Freq
        '
        Me.Freq.Location = New System.Drawing.Point(485, 97)
        Me.Freq.Name = "Freq"
        Me.Freq.Size = New System.Drawing.Size(120, 20)
        Me.Freq.TabIndex = 6
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(383, 97)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(28, 13)
        Me.Label6.TabIndex = 7
        Me.Label6.Text = "Freq"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(386, 126)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(35, 13)
        Me.Label7.TabIndex = 8
        Me.Label7.Text = "Offset"
        '
        'StartTime
        '
        Me.StartTime.Location = New System.Drawing.Point(256, 96)
        Me.StartTime.Name = "StartTime"
        Me.StartTime.Size = New System.Drawing.Size(100, 20)
        Me.StartTime.TabIndex = 9
        '
        'EndTime
        '
        Me.EndTime.Location = New System.Drawing.Point(256, 128)
        Me.EndTime.Name = "EndTime"
        Me.EndTime.Size = New System.Drawing.Size(100, 20)
        Me.EndTime.TabIndex = 10
        '
        'Offset
        '
        Me.Offset.Location = New System.Drawing.Point(485, 126)
        Me.Offset.Name = "Offset"
        Me.Offset.Size = New System.Drawing.Size(120, 20)
        Me.Offset.TabIndex = 11
        '
        'AutoStart
        '
        Me.AutoStart.AutoSize = True
        Me.AutoStart.Location = New System.Drawing.Point(157, 172)
        Me.AutoStart.Name = "AutoStart"
        Me.AutoStart.Size = New System.Drawing.Size(526, 17)
        Me.AutoStart.TabIndex = 12
        Me.AutoStart.Text = "Start App w/ Timer Running (Next Run will automatically be set to Start Time or N" & _
    "ext Time based on Offset"
        Me.AutoStart.UseVisualStyleBackColor = True
        '
        'AutoShutDown
        '
        Me.AutoShutDown.AutoSize = True
        Me.AutoShutDown.Location = New System.Drawing.Point(157, 205)
        Me.AutoShutDown.Name = "AutoShutDown"
        Me.AutoShutDown.Size = New System.Drawing.Size(373, 17)
        Me.AutoShutDown.TabIndex = 13
        Me.AutoShutDown.Text = "Shutdown App Each Collection (Turn Off if you are using Task Scheduler)"
        Me.AutoShutDown.UseVisualStyleBackColor = True
        '
        'RunOnTimer
        '
        Me.RunOnTimer.AutoSize = True
        Me.RunOnTimer.Location = New System.Drawing.Point(157, 58)
        Me.RunOnTimer.Name = "RunOnTimer"
        Me.RunOnTimer.Size = New System.Drawing.Size(384, 17)
        Me.RunOnTimer.TabIndex = 14
        Me.RunOnTimer.Text = "Run On Timer (Delay Start Enabled - If Unchecked you must Start Manually)"
        Me.RunOnTimer.UseVisualStyleBackColor = True
        '
        'FromUserName
        '
        Me.FromUserName.Location = New System.Drawing.Point(243, 273)
        Me.FromUserName.Name = "FromUserName"
        Me.FromUserName.Size = New System.Drawing.Size(362, 20)
        Me.FromUserName.TabIndex = 15
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(154, 273)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(58, 13)
        Me.Label8.TabIndex = 16
        Me.Label8.Text = "Email From"
        '
        'DistributeOnServer
        '
        Me.DistributeOnServer.AutoSize = True
        Me.DistributeOnServer.Location = New System.Drawing.Point(157, 239)
        Me.DistributeOnServer.Name = "DistributeOnServer"
        Me.DistributeOnServer.Size = New System.Drawing.Size(498, 17)
        Me.DistributeOnServer.TabIndex = 17
        Me.DistributeOnServer.Text = "Distribute / Send Emails from Server  (Check to Send Emails in Batch after All re" & _
    "ports are Generated)"
        Me.DistributeOnServer.UseVisualStyleBackColor = True
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(154, 351)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(86, 13)
        Me.Label9.TabIndex = 19
        Me.Label9.Text = "SMTP Password"
        '
        'SMTPPassword
        '
        Me.SMTPPassword.Location = New System.Drawing.Point(243, 351)
        Me.SMTPPassword.Name = "SMTPPassword"
        Me.SMTPPassword.Size = New System.Drawing.Size(362, 20)
        Me.SMTPPassword.TabIndex = 18
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(154, 325)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(62, 13)
        Me.Label10.TabIndex = 21
        Me.Label10.Text = "SMTP User"
        '
        'SMTPUserID
        '
        Me.SMTPUserID.Location = New System.Drawing.Point(243, 325)
        Me.SMTPUserID.Name = "SMTPUserID"
        Me.SMTPUserID.Size = New System.Drawing.Size(362, 20)
        Me.SMTPUserID.TabIndex = 20
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(154, 299)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(71, 13)
        Me.Label11.TabIndex = 23
        Me.Label11.Text = "SMTP Server"
        '
        'SMTPServer
        '
        Me.SMTPServer.Location = New System.Drawing.Point(243, 299)
        Me.SMTPServer.Name = "SMTPServer"
        Me.SMTPServer.Size = New System.Drawing.Size(362, 20)
        Me.SMTPServer.TabIndex = 22
        '
        'Settings
        '
        Me.ClientSize = New System.Drawing.Size(710, 517)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.SMTPServer)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.SMTPUserID)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.SMTPPassword)
        Me.Controls.Add(Me.DistributeOnServer)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.FromUserName)
        Me.Controls.Add(Me.RunOnTimer)
        Me.Controls.Add(Me.AutoShutDown)
        Me.Controls.Add(Me.AutoStart)
        Me.Controls.Add(Me.Offset)
        Me.Controls.Add(Me.EndTime)
        Me.Controls.Add(Me.StartTime)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Freq)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.WorkStationID)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Settings"
        CType(Me.Freq, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Offset, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents SendFrom As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents WorkstationName As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents WorkStationID As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Freq As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents StartTime As System.Windows.Forms.TextBox
    Friend WithEvents EndTime As System.Windows.Forms.TextBox
    Friend WithEvents Offset As System.Windows.Forms.NumericUpDown
    Friend WithEvents AutoStart As System.Windows.Forms.CheckBox
    Friend WithEvents AutoShutDown As System.Windows.Forms.CheckBox
    Friend WithEvents RunOnTimer As System.Windows.Forms.CheckBox
    Friend WithEvents FromUserName As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents DistributeOnServer As System.Windows.Forms.CheckBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents SMTPPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents SMTPUserID As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents SMTPServer As System.Windows.Forms.TextBox
    '   Friend WithEvents Close As System.Windows.Forms.Button
End Class
