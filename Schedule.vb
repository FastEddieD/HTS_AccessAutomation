Imports System
Imports System.Data
Imports System.Net.Sockets
Imports System.Net.DnsPermissionAttribute
Imports System.Security.Permissions
Imports System.Windows.Forms
Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Web.Services
Imports System.Net.Mail
Imports System.Configuration
Imports System.Threading
Imports Microsoft.Office
Imports System.Collections
Imports Microsoft.Office.Interop
Imports Microsoft.WindowsAzure
' Imports Microsoft.WindowsAzure.Storage
Imports Microsoft.WindowsAzure.StorageClient
Imports AccessAutomation.AutoReportsWCFService
Imports System.Diagnostics.Process
Imports Microsoft.Office.Interop.Access
Imports Microsoft.VisualBasic
Imports SendGrid

' Imports SendGrid.Helpers.Mail


<DnsPermissionAttribute(SecurityAction.Demand, Unrestricted:=True)> Public Class Schedule
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "
    '    Dim ardata As New DataSetLocal
    '    Dim ardata As New DataSetLocal
    Dim ardata As New DataSet
    '   Dim ar As New AutoReportsWS3.Service1
    Dim ws As New AutoReportsWCFService.ServiceClient
    Dim ARCheckComplete As Boolean = False
    Friend WithEvents RunReportsButton As System.Windows.Forms.Button
    Friend WithEvents LastCheck As System.Windows.Forms.TextBox
    Friend WithEvents RunOnTimer As System.Windows.Forms.CheckBox
    Friend WithEvents OverdueList As System.Windows.Forms.TabPage
    Friend WithEvents BGReportGenerator As System.ComponentModel.BackgroundWorker
    Friend WithEvents StartButton2 As System.Windows.Forms.ToolBarButton
    Friend WithEvents dgvOverdueJobs As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents StartButton As System.Windows.Forms.Button
    Friend WithEvents StopButton As System.Windows.Forms.Button
    Friend WithEvents SettingsButton As System.Windows.Forms.Button
    Friend WithEvents RunButton As System.Windows.Forms.Button
    Friend WithEvents RefreshButton As System.Windows.Forms.Button
    Friend WithEvents OpenCloseAccess As System.Windows.Forms.Button
    Friend WithEvents BGReportDistributor As System.ComponentModel.BackgroundWorker
    Friend WithEvents NextDistribute As System.Windows.Forms.TextBox
    Friend WithEvents Distributing As System.Windows.Forms.CheckBox
    Friend WithEvents LastDistributeCheck As System.Windows.Forms.TextBox
    Friend WithEvents DistributeButton As System.Windows.Forms.Button
    Friend WithEvents Distribute As System.Windows.Forms.TabPage
    Friend WithEvents dgvDistributionList As System.Windows.Forms.DataGridView
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents dgvMessageList As System.Windows.Forms.DataGridView
    Friend WithEvents ACCESSUSERDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LOTDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SENTBYDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents MessageListBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents BGMessenger As System.ComponentModel.BackgroundWorker
    Friend WithEvents NextMessageRun As System.Windows.Forms.TextBox
    Friend WithEvents SendingMessages As System.Windows.Forms.CheckBox
    Friend WithEvents LastMessageRun As System.Windows.Forms.TextBox
    Friend WithEvents Messenger As System.Windows.Forms.Button
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents dgvAccessAlerts As System.Windows.Forms.DataGridView
    Friend WithEvents MessageListBindingSource1 As System.Windows.Forms.BindingSource
    Friend WithEvents AlertListBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents RECIPIENTDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents QUEUENAME As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents MSGCOUNTDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LOCATION As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents OLDEST As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NEWEST As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ELASPEDMINUTESDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ISACTIVEDataGridViewCheckBoxColumn1 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents JOBIDDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TYPEDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents REPORTNAMEDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DESCRIPTIONDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FREQDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents INTERVALDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LASTRUNDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NEXTSCHEDDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CRITERIADataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CONTAINERDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents OUTPUTFORMATDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LASTARCHIVEPATHDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DISTLISTDataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SELECTEDDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents JobListBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents APPIDDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CONTAINERDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CRITERIADataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DESCRIPTIONDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DISTLISTDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FREQDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ICONDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents INTERVALDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ISACTIVEDataGridViewCheckBoxColumn As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents JOBIDDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LASTARCHIVEPATHDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LASTRUNDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NEXTSCHEDDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents OUTPUTFORMATDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents REPORTIDDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents REPORTNAMEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SELECTEDDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TYPEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents QADataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents MESSAGETYPEDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CREATEDDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ACCESSUSERIDDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DISTLISTDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SUBJECTDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TRIGGERDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DELFLAGDataGridViewCheckBoxColumn As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ACTIVEDataGridViewCheckBoxColumn As DataGridViewCheckBoxColumn
    Friend WithEvents APPLICATIONDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As DataGridViewTextBoxColumn
    Friend WithEvents ENDTIMEOFDAYDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As DataGridViewTextBoxColumn
    Friend WithEvents FROMDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents JOBDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents LASTSENTDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn5 As DataGridViewTextBoxColumn
    Friend WithEvents OUTPUTDIRECTORYDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents OUTPUTFILENAMEDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn6 As DataGridViewTextBoxColumn
    Friend WithEvents QTYDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents REPORTDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents REPORTOLDDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents STARTDATEDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents STARTTIMEOFDAYDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents JOBBindingSource As BindingSource
    Friend WithEvents DataSetLocal As DataSetLocal
    Friend WithEvents Button1 As Button
    Dim ASNCheckComplete As Boolean = False
    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()



        ' Load Startup Parameters
        LoadSettings()


        Me.StatusBar.Text = "Retrieving Report Schedule..."
        refreshdata()
        '   Me.Show()

        Timer1.Interval = 1000
        Timer1.Enabled = True
        Timer1.Start()





        '  startlistener()c
    End Sub

    Private Function GetNextCheckTime(StartTime As Date, Offset As Long, Freq As Long, EndTime As Date) As Date
        '  Add Offset to StartTime


        Dim STOD As Date
        Dim NTOD As Date
        Dim CTOD As Date
        Dim ETOD As Date
        Dim strStartTime As String
        Dim strEndTime As String


        strStartTime = CStr(Date.Today.Date) + " " + CStr(StartTime.AddMinutes(Offset))
        strEndTime = CStr(Date.Today.Date) + " " + CStr(EndTime)




        STOD = CDate(strStartTime)
        ETOD = CDate(strEndTime)
        CTOD = Now
        NTOD = STOD
        If DateDiff("n", STOD, CTOD) < 0 Then
            Return STOD
        End If
        '  Increment Forward Until you get to a Time greater than now
        Do Until NTOD >= Now
            NTOD = NTOD.AddMinutes(Freq)
        Loop
        If NTOD > ETOD Then
            Return STOD.AddDays(1)
        Else
            Return NTOD
        End If
        '  Return Next Time


    End Function


    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Timer1 As System.Timers.Timer
    Friend WithEvents NextCheck As System.Windows.Forms.TextBox
    Friend WithEvents Running As System.Windows.Forms.CheckBox
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents StatusBar As System.Windows.Forms.TextBox
    Friend WithEvents TreeView1 As System.Windows.Forms.TreeView
    '  Friend WithEvents DataSet12 As AccessAutomation.DataSet1
    '    Friend WithEvents DataSet11 As AccessAutomation.DataSet1
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents JobSchedule As System.Windows.Forms.TabPage
    Friend WithEvents ActivityLog As System.Windows.Forms.TabPage
    Friend WithEvents Users As System.Windows.Forms.TabPage
    Friend WithEvents Applications As System.Windows.Forms.TabPage
    Friend WithEvents Reports As System.Windows.Forms.TabPage
    Friend WithEvents DataGridTextBoxColumn1 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents dgUserList As System.Windows.Forms.DataGridTableStyle
    Friend WithEvents DataGridTextBoxColumn2 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn3 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents dgActivityLog As System.Windows.Forms.DataGrid
    Friend WithEvents tsHistory As System.Windows.Forms.DataGridTableStyle
    Friend WithEvents DateTimeRun As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents JobNo As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn4 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn5 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn6 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents dgUserSubscriptions As System.Windows.Forms.DataGrid
    Friend WithEvents dgUserList2 As System.Windows.Forms.DataGrid
    Friend WithEvents tsUserSubscriptions As System.Windows.Forms.DataGridTableStyle
    Friend WithEvents DataGridTextBoxColumn7 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn8 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents Online As System.Windows.Forms.CheckBox
    Friend WithEvents WorkOffline As System.Windows.Forms.CheckBox
    Friend WithEvents StartOffline As System.Windows.Forms.CheckBox
    Friend WithEvents DataSetLocal1 As AccessAutomation.DataSetLocal
    Friend WithEvents ASN As System.Windows.Forms.TabPage
    Friend WithEvents dgASNList As System.Windows.Forms.DataGrid
    Friend WithEvents EmailTest As System.Windows.Forms.Button

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Schedule))
        Dim configurationAppSettings As System.Configuration.AppSettingsReader = New System.Configuration.AppSettingsReader()
        Me.Timer1 = New System.Timers.Timer()
        Me.NextCheck = New System.Windows.Forms.TextBox()
        Me.Running = New System.Windows.Forms.CheckBox()
        Me.StatusBar = New System.Windows.Forms.TextBox()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.TreeView1 = New System.Windows.Forms.TreeView()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.JobSchedule = New System.Windows.Forms.TabPage()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.JobListBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.ActivityLog = New System.Windows.Forms.TabPage()
        Me.dgActivityLog = New System.Windows.Forms.DataGrid()
        Me.tsHistory = New System.Windows.Forms.DataGridTableStyle()
        Me.DateTimeRun = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.JobNo = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn4 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn5 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn6 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.Users = New System.Windows.Forms.TabPage()
        Me.dgUserSubscriptions = New System.Windows.Forms.DataGrid()
        Me.tsUserSubscriptions = New System.Windows.Forms.DataGridTableStyle()
        Me.DataGridTextBoxColumn7 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn8 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.dgUserList2 = New System.Windows.Forms.DataGrid()
        Me.dgUserList = New System.Windows.Forms.DataGridTableStyle()
        Me.DataGridTextBoxColumn1 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn2 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn3 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.Applications = New System.Windows.Forms.TabPage()
        Me.Reports = New System.Windows.Forms.TabPage()
        Me.DataGrid1 = New System.Windows.Forms.DataGrid()
        Me.ASN = New System.Windows.Forms.TabPage()
        Me.dgASNList = New System.Windows.Forms.DataGrid()
        Me.OverdueList = New System.Windows.Forms.TabPage()
        Me.dgvOverdueJobs = New System.Windows.Forms.DataGridView()
        Me.ACTIVEDataGridViewCheckBoxColumn = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.APPLICATIONDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ENDTIMEOFDAYDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FROMDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.JOBDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LASTSENTDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OUTPUTDIRECTORYDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OUTPUTFILENAMEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.QTYDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.REPORTDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.REPORTOLDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.STARTDATEDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.STARTTIMEOFDAYDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.JOBBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.DataSetLocal = New AccessAutomation.DataSetLocal()
        Me.Distribute = New System.Windows.Forms.TabPage()
        Me.dgvDistributionList = New System.Windows.Forms.DataGridView()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.dgvMessageList = New System.Windows.Forms.DataGridView()
        Me.MessageListBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.dgvAccessAlerts = New System.Windows.Forms.DataGridView()
        Me.QUEUENAME = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LOCATION = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OLDEST = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NEWEST = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.AlertListBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.MessageListBindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
        Me.Online = New System.Windows.Forms.CheckBox()
        Me.WorkOffline = New System.Windows.Forms.CheckBox()
        Me.StartOffline = New System.Windows.Forms.CheckBox()
        Me.EmailTest = New System.Windows.Forms.Button()
        Me.RunReportsButton = New System.Windows.Forms.Button()
        Me.LastCheck = New System.Windows.Forms.TextBox()
        Me.RunOnTimer = New System.Windows.Forms.CheckBox()
        Me.BGReportGenerator = New System.ComponentModel.BackgroundWorker()
        Me.StartButton = New System.Windows.Forms.Button()
        Me.StopButton = New System.Windows.Forms.Button()
        Me.RefreshButton = New System.Windows.Forms.Button()
        Me.RunButton = New System.Windows.Forms.Button()
        Me.SettingsButton = New System.Windows.Forms.Button()
        Me.OpenCloseAccess = New System.Windows.Forms.Button()
        Me.BGReportDistributor = New System.ComponentModel.BackgroundWorker()
        Me.Distributing = New System.Windows.Forms.CheckBox()
        Me.NextDistribute = New System.Windows.Forms.TextBox()
        Me.LastDistributeCheck = New System.Windows.Forms.TextBox()
        Me.DistributeButton = New System.Windows.Forms.Button()
        Me.BGMessenger = New System.ComponentModel.BackgroundWorker()
        Me.NextMessageRun = New System.Windows.Forms.TextBox()
        Me.SendingMessages = New System.Windows.Forms.CheckBox()
        Me.LastMessageRun = New System.Windows.Forms.TextBox()
        Me.Messenger = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        CType(Me.Timer1,System.ComponentModel.ISupportInitialize).BeginInit
        Me.TabControl1.SuspendLayout
        Me.JobSchedule.SuspendLayout
        CType(Me.DataGridView1,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.JobListBindingSource,System.ComponentModel.ISupportInitialize).BeginInit
        Me.ActivityLog.SuspendLayout
        CType(Me.dgActivityLog,System.ComponentModel.ISupportInitialize).BeginInit
        Me.Users.SuspendLayout
        CType(Me.dgUserSubscriptions,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.dgUserList2,System.ComponentModel.ISupportInitialize).BeginInit
        Me.Reports.SuspendLayout
        CType(Me.DataGrid1,System.ComponentModel.ISupportInitialize).BeginInit
        Me.ASN.SuspendLayout
        CType(Me.dgASNList,System.ComponentModel.ISupportInitialize).BeginInit
        Me.OverdueList.SuspendLayout
        CType(Me.dgvOverdueJobs,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.JOBBindingSource,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.DataSetLocal,System.ComponentModel.ISupportInitialize).BeginInit
        Me.Distribute.SuspendLayout
        CType(Me.dgvDistributionList,System.ComponentModel.ISupportInitialize).BeginInit
        Me.TabPage1.SuspendLayout
        CType(Me.dgvMessageList,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.MessageListBindingSource,System.ComponentModel.ISupportInitialize).BeginInit
        Me.TabPage2.SuspendLayout
        CType(Me.dgvAccessAlerts,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.AlertListBindingSource,System.ComponentModel.ISupportInitialize).BeginInit
        CType(Me.MessageListBindingSource1,System.ComponentModel.ISupportInitialize).BeginInit
        Me.SuspendLayout
        '
        'Timer1
        '
        Me.Timer1.Enabled = true
        Me.Timer1.Interval = 1000R
        Me.Timer1.SynchronizingObject = Me
        '
        'NextCheck
        '
        Me.NextCheck.Location = New System.Drawing.Point(1228, 207)
        Me.NextCheck.Name = "NextCheck"
        Me.NextCheck.Size = New System.Drawing.Size(160, 20)
        Me.NextCheck.TabIndex = 1
        Me.NextCheck.Text = "Next Report Check"
        '
        'Running
        '
        Me.Running.Checked = true
        Me.Running.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Running.Location = New System.Drawing.Point(1091, 209)
        Me.Running.Name = "Running"
        Me.Running.Size = New System.Drawing.Size(96, 16)
        Me.Running.TabIndex = 2
        Me.Running.Text = "Running"
        '
        'StatusBar
        '
        Me.StatusBar.BackColor = System.Drawing.SystemColors.Control
        Me.StatusBar.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
        Me.StatusBar.Location = New System.Drawing.Point(53, 71)
        Me.StatusBar.Name = "StatusBar"
        Me.StatusBar.Size = New System.Drawing.Size(1063, 22)
        Me.StatusBar.TabIndex = 9
        Me.StatusBar.Text = "StatusBar"
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"),System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "")
        Me.ImageList1.Images.SetKeyName(1, "")
        '
        'TreeView1
        '
        Me.TreeView1.Dock = System.Windows.Forms.DockStyle.Left
        Me.TreeView1.Location = New System.Drawing.Point(0, 0)
        Me.TreeView1.Name = "TreeView1"
        Me.TreeView1.Size = New System.Drawing.Size(35, 558)
        Me.TreeView1.TabIndex = 14
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.JobSchedule)
        Me.TabControl1.Controls.Add(Me.ActivityLog)
        Me.TabControl1.Controls.Add(Me.Users)
        Me.TabControl1.Controls.Add(Me.Applications)
        Me.TabControl1.Controls.Add(Me.Reports)
        Me.TabControl1.Controls.Add(Me.ASN)
        Me.TabControl1.Controls.Add(Me.OverdueList)
        Me.TabControl1.Controls.Add(Me.Distribute)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(53, 110)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(960, 382)
        Me.TabControl1.TabIndex = 20
        '
        'JobSchedule
        '
        Me.JobSchedule.Controls.Add(Me.DataGridView1)
        Me.JobSchedule.Location = New System.Drawing.Point(4, 22)
        Me.JobSchedule.Name = "JobSchedule"
        Me.JobSchedule.Size = New System.Drawing.Size(952, 356)
        Me.JobSchedule.TabIndex = 0
        Me.JobSchedule.Text = "Job Schedule"
        '
        'DataGridView1
        '
        Me.DataGridView1.AutoGenerateColumns = false
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.DataSource = Me.JobListBindingSource
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(0, 0)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(952, 356)
        Me.DataGridView1.TabIndex = 0
        '
        'ActivityLog
        '
        Me.ActivityLog.Controls.Add(Me.dgActivityLog)
        Me.ActivityLog.Location = New System.Drawing.Point(4, 22)
        Me.ActivityLog.Name = "ActivityLog"
        Me.ActivityLog.Size = New System.Drawing.Size(952, 356)
        Me.ActivityLog.TabIndex = 1
        Me.ActivityLog.Text = "Activity Log"
        '
        'dgActivityLog
        '
        Me.dgActivityLog.CaptionText = "Activity Log"
        Me.dgActivityLog.DataMember = ""
        Me.dgActivityLog.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dgActivityLog.Location = New System.Drawing.Point(-4, 11)
        Me.dgActivityLog.Name = "dgActivityLog"
        Me.dgActivityLog.Size = New System.Drawing.Size(694, 325)
        Me.dgActivityLog.TabIndex = 20
        Me.dgActivityLog.TableStyles.AddRange(New System.Windows.Forms.DataGridTableStyle() {Me.tsHistory})
        '
        'tsHistory
        '
        Me.tsHistory.DataGrid = Me.dgActivityLog
        Me.tsHistory.GridColumnStyles.AddRange(New System.Windows.Forms.DataGridColumnStyle() {Me.DateTimeRun, Me.JobNo, Me.DataGridTextBoxColumn4, Me.DataGridTextBoxColumn5, Me.DataGridTextBoxColumn6})
        Me.tsHistory.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.tsHistory.MappingName = "HISTORY"
        '
        'DateTimeRun
        '
        Me.DateTimeRun.Format = "g"
        Me.DateTimeRun.FormatInfo = Nothing
        Me.DateTimeRun.HeaderText = "Date / Time Run"
        Me.DateTimeRun.MappingName = "DATETIME_RUN"
        Me.DateTimeRun.Width = 125
        '
        'JobNo
        '
        Me.JobNo.Format = ""
        Me.JobNo.FormatInfo = Nothing
        Me.JobNo.HeaderText = "Job"
        Me.JobNo.MappingName = "JOB"
        Me.JobNo.Width = 50
        '
        'DataGridTextBoxColumn4
        '
        Me.DataGridTextBoxColumn4.Format = ""
        Me.DataGridTextBoxColumn4.FormatInfo = Nothing
        Me.DataGridTextBoxColumn4.HeaderText = "Action"
        Me.DataGridTextBoxColumn4.MappingName = "ACTION"
        Me.DataGridTextBoxColumn4.Width = 250
        '
        'DataGridTextBoxColumn5
        '
        Me.DataGridTextBoxColumn5.Format = ""
        Me.DataGridTextBoxColumn5.FormatInfo = Nothing
        Me.DataGridTextBoxColumn5.HeaderText = "Distribution List"
        Me.DataGridTextBoxColumn5.MappingName = "DISTRIBUTION_LIST"
        Me.DataGridTextBoxColumn5.Width = 250
        '
        'DataGridTextBoxColumn6
        '
        Me.DataGridTextBoxColumn6.Format = "g"
        Me.DataGridTextBoxColumn6.FormatInfo = Nothing
        Me.DataGridTextBoxColumn6.HeaderText = "Scheduled"
        Me.DataGridTextBoxColumn6.MappingName = "DATETIME_SCHED"
        Me.DataGridTextBoxColumn6.Width = 125
        '
        'Users
        '
        Me.Users.Controls.Add(Me.dgUserSubscriptions)
        Me.Users.Controls.Add(Me.dgUserList2)
        Me.Users.Location = New System.Drawing.Point(4, 22)
        Me.Users.Name = "Users"
        Me.Users.Size = New System.Drawing.Size(952, 356)
        Me.Users.TabIndex = 2
        Me.Users.Text = "Users / Subsriptions"
        '
        'dgUserSubscriptions
        '
        Me.dgUserSubscriptions.DataMember = ""
        Me.dgUserSubscriptions.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dgUserSubscriptions.Location = New System.Drawing.Point(0, 176)
        Me.dgUserSubscriptions.Name = "dgUserSubscriptions"
        Me.dgUserSubscriptions.Size = New System.Drawing.Size(242, 83)
        Me.dgUserSubscriptions.TabIndex = 2
        Me.dgUserSubscriptions.TableStyles.AddRange(New System.Windows.Forms.DataGridTableStyle() {Me.tsUserSubscriptions})
        '
        'tsUserSubscriptions
        '
        Me.tsUserSubscriptions.DataGrid = Me.dgUserSubscriptions
        Me.tsUserSubscriptions.GridColumnStyles.AddRange(New System.Windows.Forms.DataGridColumnStyle() {Me.DataGridTextBoxColumn7, Me.DataGridTextBoxColumn8})
        Me.tsUserSubscriptions.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.tsUserSubscriptions.MappingName = "DISTRIBUTION_LIST"
        '
        'DataGridTextBoxColumn7
        '
        Me.DataGridTextBoxColumn7.Format = ""
        Me.DataGridTextBoxColumn7.FormatInfo = Nothing
        Me.DataGridTextBoxColumn7.HeaderText = "Job"
        Me.DataGridTextBoxColumn7.MappingName = "JOB"
        Me.DataGridTextBoxColumn7.Width = 75
        '
        'DataGridTextBoxColumn8
        '
        Me.DataGridTextBoxColumn8.Format = ""
        Me.DataGridTextBoxColumn8.FormatInfo = Nothing
        Me.DataGridTextBoxColumn8.MappingName = "ACTIVE"
        Me.DataGridTextBoxColumn8.Width = 75
        '
        'dgUserList2
        '
        Me.dgUserList2.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.dgUserList2.DataMember = ""
        Me.dgUserList2.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dgUserList2.Location = New System.Drawing.Point(12, 0)
        Me.dgUserList2.Name = "dgUserList2"
        Me.dgUserList2.Size = New System.Drawing.Size(266, 67)
        Me.dgUserList2.TabIndex = 0
        Me.dgUserList2.TableStyles.AddRange(New System.Windows.Forms.DataGridTableStyle() {Me.dgUserList})
        '
        'dgUserList
        '
        Me.dgUserList.DataGrid = Me.dgUserList2
        Me.dgUserList.GridColumnStyles.AddRange(New System.Windows.Forms.DataGridColumnStyle() {Me.DataGridTextBoxColumn1, Me.DataGridTextBoxColumn2, Me.DataGridTextBoxColumn3})
        Me.dgUserList.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dgUserList.MappingName = "USERS"
        '
        'DataGridTextBoxColumn1
        '
        Me.DataGridTextBoxColumn1.Format = ">"
        Me.DataGridTextBoxColumn1.FormatInfo = Nothing
        Me.DataGridTextBoxColumn1.HeaderText = "Contact / Name"
        Me.DataGridTextBoxColumn1.MappingName = "CONTACT"
        Me.DataGridTextBoxColumn1.Width = 200
        '
        'DataGridTextBoxColumn2
        '
        Me.DataGridTextBoxColumn2.Format = ""
        Me.DataGridTextBoxColumn2.FormatInfo = Nothing
        Me.DataGridTextBoxColumn2.HeaderText = "Email Address"
        Me.DataGridTextBoxColumn2.MappingName = "EMAIL"
        Me.DataGridTextBoxColumn2.Width = 300
        '
        'DataGridTextBoxColumn3
        '
        Me.DataGridTextBoxColumn3.Format = ""
        Me.DataGridTextBoxColumn3.FormatInfo = Nothing
        Me.DataGridTextBoxColumn3.HeaderText = "Active"
        Me.DataGridTextBoxColumn3.MappingName = "ACTIVE"
        Me.DataGridTextBoxColumn3.Width = 75
        '
        'Applications
        '
        Me.Applications.Location = New System.Drawing.Point(4, 22)
        Me.Applications.Name = "Applications"
        Me.Applications.Size = New System.Drawing.Size(952, 356)
        Me.Applications.TabIndex = 3
        Me.Applications.Text = "Applications / Reports"
        '
        'Reports
        '
        Me.Reports.Controls.Add(Me.DataGrid1)
        Me.Reports.Location = New System.Drawing.Point(4, 22)
        Me.Reports.Name = "Reports"
        Me.Reports.Size = New System.Drawing.Size(952, 356)
        Me.Reports.TabIndex = 4
        Me.Reports.Text = "Reports / Schedule"
        '
        'DataGrid1
        '
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Location = New System.Drawing.Point(40, 96)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.Size = New System.Drawing.Size(225, 96)
        Me.DataGrid1.TabIndex = 0
        '
        'ASN
        '
        Me.ASN.Controls.Add(Me.dgASNList)
        Me.ASN.Location = New System.Drawing.Point(4, 22)
        Me.ASN.Name = "ASN"
        Me.ASN.Size = New System.Drawing.Size(952, 356)
        Me.ASN.TabIndex = 5
        Me.ASN.Text = "ASN List"
        '
        'dgASNList
        '
        Me.dgASNList.DataMember = ""
        Me.dgASNList.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dgASNList.Location = New System.Drawing.Point(82, 8)
        Me.dgASNList.Name = "dgASNList"
        Me.dgASNList.Size = New System.Drawing.Size(290, 129)
        Me.dgASNList.TabIndex = 0
        '
        'OverdueList
        '
        Me.OverdueList.Controls.Add(Me.dgvOverdueJobs)
        Me.OverdueList.Location = New System.Drawing.Point(4, 22)
        Me.OverdueList.Name = "OverdueList"
        Me.OverdueList.Size = New System.Drawing.Size(952, 356)
        Me.OverdueList.TabIndex = 6
        Me.OverdueList.Text = "OverdueList"
        Me.OverdueList.UseVisualStyleBackColor = true
        '
        'dgvOverdueJobs
        '
        Me.dgvOverdueJobs.AutoGenerateColumns = false
        Me.dgvOverdueJobs.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvOverdueJobs.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ACTIVEDataGridViewCheckBoxColumn, Me.APPLICATIONDataGridViewTextBoxColumn, Me.DataGridViewTextBoxColumn2, Me.DataGridViewTextBoxColumn3, Me.ENDTIMEOFDAYDataGridViewTextBoxColumn, Me.DataGridViewTextBoxColumn4, Me.FROMDataGridViewTextBoxColumn, Me.JOBDataGridViewTextBoxColumn, Me.LASTSENTDataGridViewTextBoxColumn, Me.DataGridViewTextBoxColumn5, Me.OUTPUTDIRECTORYDataGridViewTextBoxColumn, Me.OUTPUTFILENAMEDataGridViewTextBoxColumn, Me.DataGridViewTextBoxColumn6, Me.QTYDataGridViewTextBoxColumn, Me.REPORTDataGridViewTextBoxColumn, Me.REPORTOLDDataGridViewTextBoxColumn, Me.STARTDATEDataGridViewTextBoxColumn, Me.STARTTIMEOFDAYDataGridViewTextBoxColumn})
        Me.dgvOverdueJobs.DataSource = Me.JOBBindingSource
        Me.dgvOverdueJobs.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvOverdueJobs.Location = New System.Drawing.Point(0, 0)
        Me.dgvOverdueJobs.Name = "dgvOverdueJobs"
        Me.dgvOverdueJobs.Size = New System.Drawing.Size(952, 356)
        Me.dgvOverdueJobs.TabIndex = 0
        '
        'ACTIVEDataGridViewCheckBoxColumn
        '
        Me.ACTIVEDataGridViewCheckBoxColumn.DataPropertyName = "ACTIVE"
        Me.ACTIVEDataGridViewCheckBoxColumn.HeaderText = "ACTIVE"
        Me.ACTIVEDataGridViewCheckBoxColumn.Name = "ACTIVEDataGridViewCheckBoxColumn"
        '
        'APPLICATIONDataGridViewTextBoxColumn
        '
        Me.APPLICATIONDataGridViewTextBoxColumn.DataPropertyName = "APPLICATION"
        Me.APPLICATIONDataGridViewTextBoxColumn.HeaderText = "APPLICATION"
        Me.APPLICATIONDataGridViewTextBoxColumn.Name = "APPLICATIONDataGridViewTextBoxColumn"
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.DataPropertyName = "CRITERIA"
        Me.DataGridViewTextBoxColumn2.HeaderText = "CRITERIA"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.DataPropertyName = "DESCRIPTION"
        Me.DataGridViewTextBoxColumn3.HeaderText = "DESCRIPTION"
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        '
        'ENDTIMEOFDAYDataGridViewTextBoxColumn
        '
        Me.ENDTIMEOFDAYDataGridViewTextBoxColumn.DataPropertyName = "END_TIME_OF_DAY"
        Me.ENDTIMEOFDAYDataGridViewTextBoxColumn.HeaderText = "END_TIME_OF_DAY"
        Me.ENDTIMEOFDAYDataGridViewTextBoxColumn.Name = "ENDTIMEOFDAYDataGridViewTextBoxColumn"
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.DataPropertyName = "FREQ"
        Me.DataGridViewTextBoxColumn4.HeaderText = "FREQ"
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        '
        'FROMDataGridViewTextBoxColumn
        '
        Me.FROMDataGridViewTextBoxColumn.DataPropertyName = "FROM"
        Me.FROMDataGridViewTextBoxColumn.HeaderText = "FROM"
        Me.FROMDataGridViewTextBoxColumn.Name = "FROMDataGridViewTextBoxColumn"
        '
        'JOBDataGridViewTextBoxColumn
        '
        Me.JOBDataGridViewTextBoxColumn.DataPropertyName = "JOB"
        Me.JOBDataGridViewTextBoxColumn.HeaderText = "JOB"
        Me.JOBDataGridViewTextBoxColumn.Name = "JOBDataGridViewTextBoxColumn"
        '
        'LASTSENTDataGridViewTextBoxColumn
        '
        Me.LASTSENTDataGridViewTextBoxColumn.DataPropertyName = "LAST_SENT"
        Me.LASTSENTDataGridViewTextBoxColumn.HeaderText = "LAST_SENT"
        Me.LASTSENTDataGridViewTextBoxColumn.Name = "LASTSENTDataGridViewTextBoxColumn"
        '
        'DataGridViewTextBoxColumn5
        '
        Me.DataGridViewTextBoxColumn5.DataPropertyName = "NEXT_SCHED"
        Me.DataGridViewTextBoxColumn5.HeaderText = "NEXT_SCHED"
        Me.DataGridViewTextBoxColumn5.Name = "DataGridViewTextBoxColumn5"
        '
        'OUTPUTDIRECTORYDataGridViewTextBoxColumn
        '
        Me.OUTPUTDIRECTORYDataGridViewTextBoxColumn.DataPropertyName = "OUTPUT_DIRECTORY"
        Me.OUTPUTDIRECTORYDataGridViewTextBoxColumn.HeaderText = "OUTPUT_DIRECTORY"
        Me.OUTPUTDIRECTORYDataGridViewTextBoxColumn.Name = "OUTPUTDIRECTORYDataGridViewTextBoxColumn"
        '
        'OUTPUTFILENAMEDataGridViewTextBoxColumn
        '
        Me.OUTPUTFILENAMEDataGridViewTextBoxColumn.DataPropertyName = "OUTPUT_FILENAME"
        Me.OUTPUTFILENAMEDataGridViewTextBoxColumn.HeaderText = "OUTPUT_FILENAME"
        Me.OUTPUTFILENAMEDataGridViewTextBoxColumn.Name = "OUTPUTFILENAMEDataGridViewTextBoxColumn"
        '
        'DataGridViewTextBoxColumn6
        '
        Me.DataGridViewTextBoxColumn6.DataPropertyName = "OUTPUT_FORMAT"
        Me.DataGridViewTextBoxColumn6.HeaderText = "OUTPUT_FORMAT"
        Me.DataGridViewTextBoxColumn6.Name = "DataGridViewTextBoxColumn6"
        '
        'QTYDataGridViewTextBoxColumn
        '
        Me.QTYDataGridViewTextBoxColumn.DataPropertyName = "QTY"
        Me.QTYDataGridViewTextBoxColumn.HeaderText = "QTY"
        Me.QTYDataGridViewTextBoxColumn.Name = "QTYDataGridViewTextBoxColumn"
        '
        'REPORTDataGridViewTextBoxColumn
        '
        Me.REPORTDataGridViewTextBoxColumn.DataPropertyName = "REPORT"
        Me.REPORTDataGridViewTextBoxColumn.HeaderText = "REPORT"
        Me.REPORTDataGridViewTextBoxColumn.Name = "REPORTDataGridViewTextBoxColumn"
        '
        'REPORTOLDDataGridViewTextBoxColumn
        '
        Me.REPORTOLDDataGridViewTextBoxColumn.DataPropertyName = "REPORT_OLD"
        Me.REPORTOLDDataGridViewTextBoxColumn.HeaderText = "REPORT_OLD"
        Me.REPORTOLDDataGridViewTextBoxColumn.Name = "REPORTOLDDataGridViewTextBoxColumn"
        '
        'STARTDATEDataGridViewTextBoxColumn
        '
        Me.STARTDATEDataGridViewTextBoxColumn.DataPropertyName = "START_DATE"
        Me.STARTDATEDataGridViewTextBoxColumn.HeaderText = "START_DATE"
        Me.STARTDATEDataGridViewTextBoxColumn.Name = "STARTDATEDataGridViewTextBoxColumn"
        '
        'STARTTIMEOFDAYDataGridViewTextBoxColumn
        '
        Me.STARTTIMEOFDAYDataGridViewTextBoxColumn.DataPropertyName = "START_TIME_OF_DAY"
        Me.STARTTIMEOFDAYDataGridViewTextBoxColumn.HeaderText = "START_TIME_OF_DAY"
        Me.STARTTIMEOFDAYDataGridViewTextBoxColumn.Name = "STARTTIMEOFDAYDataGridViewTextBoxColumn"
        '
        'JOBBindingSource
        '
        Me.JOBBindingSource.DataMember = "JOB"
        Me.JOBBindingSource.DataSource = Me.DataSetLocal
        '
        'DataSetLocal
        '
        Me.DataSetLocal.DataSetName = "DataSetLocal"
        Me.DataSetLocal.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'Distribute
        '
        Me.Distribute.Controls.Add(Me.dgvDistributionList)
        Me.Distribute.Location = New System.Drawing.Point(4, 22)
        Me.Distribute.Name = "Distribute"
        Me.Distribute.Padding = New System.Windows.Forms.Padding(3)
        Me.Distribute.Size = New System.Drawing.Size(952, 356)
        Me.Distribute.TabIndex = 7
        Me.Distribute.Text = "Distribute List"
        Me.Distribute.UseVisualStyleBackColor = true
        '
        'dgvDistributionList
        '
        Me.dgvDistributionList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDistributionList.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1})
        Me.dgvDistributionList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvDistributionList.Location = New System.Drawing.Point(3, 3)
        Me.dgvDistributionList.Name = "dgvDistributionList"
        Me.dgvDistributionList.Size = New System.Drawing.Size(946, 350)
        Me.dgvDistributionList.TabIndex = 0
        '
        'Column1
        '
        Me.Column1.HeaderText = "Column1"
        Me.Column1.Name = "Column1"
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgvMessageList)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(952, 356)
        Me.TabPage1.TabIndex = 8
        Me.TabPage1.Text = "EmailMessages"
        Me.TabPage1.UseVisualStyleBackColor = true
        '
        'dgvMessageList
        '
        Me.dgvMessageList.AutoGenerateColumns = false
        Me.dgvMessageList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvMessageList.DataSource = Me.MessageListBindingSource
        Me.dgvMessageList.Location = New System.Drawing.Point(12, 22)
        Me.dgvMessageList.Name = "dgvMessageList"
        Me.dgvMessageList.Size = New System.Drawing.Size(918, 313)
        Me.dgvMessageList.TabIndex = 0
        '
        'MessageListBindingSource
        '
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.dgvAccessAlerts)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(952, 356)
        Me.TabPage2.TabIndex = 9
        Me.TabPage2.Text = "AccessAlerts"
        Me.TabPage2.UseVisualStyleBackColor = true
        '
        'dgvAccessAlerts
        '
        Me.dgvAccessAlerts.AutoGenerateColumns = false
        Me.dgvAccessAlerts.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvAccessAlerts.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.QUEUENAME, Me.LOCATION, Me.OLDEST, Me.NEWEST})
        Me.dgvAccessAlerts.DataSource = Me.AlertListBindingSource
        Me.dgvAccessAlerts.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvAccessAlerts.Location = New System.Drawing.Point(3, 3)
        Me.dgvAccessAlerts.Name = "dgvAccessAlerts"
        Me.dgvAccessAlerts.Size = New System.Drawing.Size(946, 350)
        Me.dgvAccessAlerts.TabIndex = 0
        '
        'QUEUENAME
        '
        Me.QUEUENAME.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.QUEUENAME.DataPropertyName = "QUEUENAME"
        Me.QUEUENAME.HeaderText = "QUEUENAME"
        Me.QUEUENAME.Name = "QUEUENAME"
        '
        'LOCATION
        '
        Me.LOCATION.DataPropertyName = "LOCATION"
        Me.LOCATION.HeaderText = "LOCATION"
        Me.LOCATION.Name = "LOCATION"
        '
        'OLDEST
        '
        Me.OLDEST.DataPropertyName = "OLDEST"
        Me.OLDEST.HeaderText = "OLDEST"
        Me.OLDEST.Name = "OLDEST"
        Me.OLDEST.Width = 150
        '
        'NEWEST
        '
        Me.NEWEST.DataPropertyName = "NEWEST"
        Me.NEWEST.HeaderText = "NEWEST"
        Me.NEWEST.Name = "NEWEST"
        Me.NEWEST.Width = 150
        '
        'Online
        '
        Me.Online.Location = New System.Drawing.Point(1137, 176)
        Me.Online.Name = "Online"
        Me.Online.Size = New System.Drawing.Size(96, 16)
        Me.Online.TabIndex = 21
        Me.Online.Text = "Online"
        '
        'WorkOffline
        '
        Me.WorkOffline.Location = New System.Drawing.Point(1137, 132)
        Me.WorkOffline.Name = "WorkOffline"
        Me.WorkOffline.Size = New System.Drawing.Size(104, 16)
        Me.WorkOffline.TabIndex = 22
        Me.WorkOffline.Text = "Work Offline"
        '
        'StartOffline
        '
        Me.StartOffline.Location = New System.Drawing.Point(1137, 154)
        Me.StartOffline.Name = "StartOffline"
        Me.StartOffline.Size = New System.Drawing.Size(104, 16)
        Me.StartOffline.TabIndex = 23
        Me.StartOffline.Text = "Start Offline"
        '
        'EmailTest
        '
        Me.EmailTest.Location = New System.Drawing.Point(1140, 317)
        Me.EmailTest.Name = "EmailTest"
        Me.EmailTest.Size = New System.Drawing.Size(75, 23)
        Me.EmailTest.TabIndex = 25
        Me.EmailTest.Text = "Test Email"
        '
        'RunReportsButton
        '
        Me.RunReportsButton.Location = New System.Drawing.Point(1140, 346)
        Me.RunReportsButton.Name = "RunReportsButton"
        Me.RunReportsButton.Size = New System.Drawing.Size(75, 23)
        Me.RunReportsButton.TabIndex = 27
        Me.RunReportsButton.Text = "Run Reports"
        Me.RunReportsButton.UseVisualStyleBackColor = true
        '
        'LastCheck
        '
        Me.LastCheck.Location = New System.Drawing.Point(1231, 395)
        Me.LastCheck.Name = "LastCheck"
        Me.LastCheck.Size = New System.Drawing.Size(160, 20)
        Me.LastCheck.TabIndex = 28
        '
        'RunOnTimer
        '
        Me.RunOnTimer.Location = New System.Drawing.Point(1137, 110)
        Me.RunOnTimer.Name = "RunOnTimer"
        Me.RunOnTimer.Size = New System.Drawing.Size(104, 16)
        Me.RunOnTimer.TabIndex = 29
        Me.RunOnTimer.Text = "Start Offline"
        '
        'BGReportGenerator
        '
        Me.BGReportGenerator.WorkerReportsProgress = true
        Me.BGReportGenerator.WorkerSupportsCancellation = true
        '
        'StartButton
        '
        Me.StartButton.Location = New System.Drawing.Point(18, 28)
        Me.StartButton.Name = "StartButton"
        Me.StartButton.Size = New System.Drawing.Size(75, 23)
        Me.StartButton.TabIndex = 30
        Me.StartButton.Text = "Start"
        Me.StartButton.UseVisualStyleBackColor = true
        '
        'StopButton
        '
        Me.StopButton.Location = New System.Drawing.Point(89, 28)
        Me.StopButton.Name = "StopButton"
        Me.StopButton.Size = New System.Drawing.Size(75, 23)
        Me.StopButton.TabIndex = 31
        Me.StopButton.Text = "Stop"
        Me.StopButton.UseVisualStyleBackColor = true
        '
        'RefreshButton
        '
        Me.RefreshButton.Location = New System.Drawing.Point(162, 28)
        Me.RefreshButton.Name = "RefreshButton"
        Me.RefreshButton.Size = New System.Drawing.Size(75, 23)
        Me.RefreshButton.TabIndex = 32
        Me.RefreshButton.Text = "Refresh"
        Me.RefreshButton.UseVisualStyleBackColor = true
        '
        'RunButton
        '
        Me.RunButton.Location = New System.Drawing.Point(234, 28)
        Me.RunButton.Name = "RunButton"
        Me.RunButton.Size = New System.Drawing.Size(75, 23)
        Me.RunButton.TabIndex = 33
        Me.RunButton.Text = "Run Now"
        Me.RunButton.UseVisualStyleBackColor = true
        '
        'SettingsButton
        '
        Me.SettingsButton.Location = New System.Drawing.Point(594, 28)
        Me.SettingsButton.Name = "SettingsButton"
        Me.SettingsButton.Size = New System.Drawing.Size(75, 23)
        Me.SettingsButton.TabIndex = 34
        Me.SettingsButton.Text = "Settings"
        Me.SettingsButton.UseVisualStyleBackColor = true
        '
        'OpenCloseAccess
        '
        Me.OpenCloseAccess.Location = New System.Drawing.Point(710, 19)
        Me.OpenCloseAccess.Name = "OpenCloseAccess"
        Me.OpenCloseAccess.Size = New System.Drawing.Size(75, 23)
        Me.OpenCloseAccess.TabIndex = 35
        Me.OpenCloseAccess.Text = "Button2.1"
        Me.OpenCloseAccess.UseVisualStyleBackColor = true
        '
        'BGReportDistributor
        '
        Me.BGReportDistributor.WorkerSupportsCancellation = true
        '
        'Distributing
        '
        Me.Distributing.AutoSize = true
        Me.Distributing.Checked = true
        Me.Distributing.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Distributing.Location = New System.Drawing.Point(1091, 232)
        Me.Distributing.Name = "Distributing"
        Me.Distributing.Size = New System.Drawing.Size(78, 17)
        Me.Distributing.TabIndex = 36
        Me.Distributing.Text = "Distributing"
        Me.Distributing.UseVisualStyleBackColor = true
        '
        'NextDistribute
        '
        Me.NextDistribute.Location = New System.Drawing.Point(1228, 227)
        Me.NextDistribute.Name = "NextDistribute"
        Me.NextDistribute.Size = New System.Drawing.Size(160, 20)
        Me.NextDistribute.TabIndex = 37
        Me.NextDistribute.Text = "Next Distribution"
        '
        'LastDistributeCheck
        '
        Me.LastDistributeCheck.Location = New System.Drawing.Point(1231, 421)
        Me.LastDistributeCheck.Name = "LastDistributeCheck"
        Me.LastDistributeCheck.Size = New System.Drawing.Size(160, 20)
        Me.LastDistributeCheck.TabIndex = 38
        '
        'DistributeButton
        '
        Me.DistributeButton.Location = New System.Drawing.Point(297, 28)
        Me.DistributeButton.Name = "DistributeButton"
        Me.DistributeButton.Size = New System.Drawing.Size(75, 23)
        Me.DistributeButton.TabIndex = 39
        Me.DistributeButton.Text = "Distribute Now"
        Me.DistributeButton.UseVisualStyleBackColor = true
        '
        'BGMessenger
        '
        Me.BGMessenger.WorkerSupportsCancellation = true
        '
        'NextMessageRun
        '
        Me.NextMessageRun.Location = New System.Drawing.Point(1228, 253)
        Me.NextMessageRun.Name = "NextMessageRun"
        Me.NextMessageRun.Size = New System.Drawing.Size(160, 20)
        Me.NextMessageRun.TabIndex = 41
        '
        'SendingMessages
        '
        Me.SendingMessages.AutoSize = true
        Me.SendingMessages.Location = New System.Drawing.Point(1091, 258)
        Me.SendingMessages.Name = "SendingMessages"
        Me.SendingMessages.Size = New System.Drawing.Size(116, 17)
        Me.SendingMessages.TabIndex = 40
        Me.SendingMessages.Text = "Sending Messages"
        Me.SendingMessages.UseVisualStyleBackColor = true
        '
        'LastMessageRun
        '
        Me.LastMessageRun.Location = New System.Drawing.Point(1231, 447)
        Me.LastMessageRun.Name = "LastMessageRun"
        Me.LastMessageRun.Size = New System.Drawing.Size(160, 20)
        Me.LastMessageRun.TabIndex = 42
        '
        'Messenger
        '
        Me.Messenger.Location = New System.Drawing.Point(368, 28)
        Me.Messenger.Name = "Messenger"
        Me.Messenger.Size = New System.Drawing.Size(75, 23)
        Me.Messenger.TabIndex = 43
        Me.Messenger.Text = "Messenger"
        Me.Messenger.UseVisualStyleBackColor = true
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(933, 18)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 44
        Me.Button1.Text = "Sendgrid Test Email"
        Me.Button1.UseVisualStyleBackColor = true
        '
        'Schedule
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1454, 558)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Messenger)
        Me.Controls.Add(Me.LastMessageRun)
        Me.Controls.Add(Me.NextMessageRun)
        Me.Controls.Add(Me.SendingMessages)
        Me.Controls.Add(Me.DistributeButton)
        Me.Controls.Add(Me.LastDistributeCheck)
        Me.Controls.Add(Me.NextDistribute)
        Me.Controls.Add(Me.Distributing)
        Me.Controls.Add(Me.OpenCloseAccess)
        Me.Controls.Add(Me.SettingsButton)
        Me.Controls.Add(Me.RunButton)
        Me.Controls.Add(Me.RefreshButton)
        Me.Controls.Add(Me.StopButton)
        Me.Controls.Add(Me.StartButton)
        Me.Controls.Add(Me.RunOnTimer)
        Me.Controls.Add(Me.LastCheck)
        Me.Controls.Add(Me.RunReportsButton)
        Me.Controls.Add(Me.EmailTest)
        Me.Controls.Add(Me.StartOffline)
        Me.Controls.Add(Me.WorkOffline)
        Me.Controls.Add(Me.Online)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.StatusBar)
        Me.Controls.Add(Me.Running)
        Me.Controls.Add(Me.NextCheck)
        Me.Controls.Add(Me.TreeView1)
        Me.Name = "Schedule"
        Me.Text = "Schedule"
        Me.TopMost = CType(configurationAppSettings.GetValue("Schedule.TopMost", GetType(Boolean)),Boolean)
        CType(Me.Timer1,System.ComponentModel.ISupportInitialize).EndInit
        Me.TabControl1.ResumeLayout(false)
        Me.JobSchedule.ResumeLayout(false)
        CType(Me.DataGridView1,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.JobListBindingSource,System.ComponentModel.ISupportInitialize).EndInit
        Me.ActivityLog.ResumeLayout(false)
        CType(Me.dgActivityLog,System.ComponentModel.ISupportInitialize).EndInit
        Me.Users.ResumeLayout(false)
        CType(Me.dgUserSubscriptions,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.dgUserList2,System.ComponentModel.ISupportInitialize).EndInit
        Me.Reports.ResumeLayout(false)
        CType(Me.DataGrid1,System.ComponentModel.ISupportInitialize).EndInit
        Me.ASN.ResumeLayout(false)
        CType(Me.dgASNList,System.ComponentModel.ISupportInitialize).EndInit
        Me.OverdueList.ResumeLayout(false)
        CType(Me.dgvOverdueJobs,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.JOBBindingSource,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.DataSetLocal,System.ComponentModel.ISupportInitialize).EndInit
        Me.Distribute.ResumeLayout(false)
        CType(Me.dgvDistributionList,System.ComponentModel.ISupportInitialize).EndInit
        Me.TabPage1.ResumeLayout(false)
        CType(Me.dgvMessageList,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.MessageListBindingSource,System.ComponentModel.ISupportInitialize).EndInit
        Me.TabPage2.ResumeLayout(false)
        CType(Me.dgvAccessAlerts,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.AlertListBindingSource,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.MessageListBindingSource1,System.ComponentModel.ISupportInitialize).EndInit
        Me.ResumeLayout(false)
        Me.PerformLayout

End Sub

#End Region
    Dim StartTime As Date = CDate(My.Settings.StartTime)
    Dim Offset As Long = CInt(My.Settings.Offset)
    Dim Freq As Long = CInt(My.Settings.Freq)
    Dim DistributeFreq As Long = CInt(My.Settings.DistributeFreq)
    Dim DistributeOffset As Long = CInt(My.Settings.DistributeOffset)
    Dim EndTime As Date = CDate(My.Settings.EndTime)
    Dim MaxCount As Long = 15
    Dim SinceDate As Date = DateAdd("d", -2, DateTime.Now)
    
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

 '       Dim client = New SendGridClient("Jm44j-Q2RQqzxnsBgXFbDg")
 '       Dim from = New EmailAddress("autoreports@htsmi.com")
 '       Dim subject = "Test Email from SendGrid using api"
 '       Dim sto = New EmailAddress("edolikian@ssitroy.com",)
        Dim transportWeb As SendGrid.Web
     '   transportWeb = New SendGrid.Web("Jm44j-Q2RQqzxnsBgXFbDg")
        Dim myMessage As SendGridMessage
        myMessage = New SendGridMessage()
        myMessage.From = New MailAddress("autoreports@htsmi.com")
        myMessage.AddTo("autoreports@ssitroy.com")
        myMessage.Subject = "Testing SendGrid"
        myMessage.Text = "Hello World plain text"
        myMessage.Html = "<p>Hellow World!<p>"

        Dim credentials As NetworkCredential
        credentials = New NetworkCredential("apikey", "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")
      '  credentials = New NetworkCredential("edolikian","HtsmiX302")
        transportWeb = New Web(credentials)
        transportWeb.Deliver(myMessage)



    '    Dim testmsg As SendGrid.SendGridMessage() = New SendGrid.SendGridMessage()
    '    testmsg
    End Sub



    Delegate Sub SetUIText_Delegate(ByVal [ControlName] As System.Windows.Forms.TextBox, ByVal [text] As String)
    Private Sub SetUIText_ThreadSafe(ByVal [ControlName] As Windows.Forms.TextBox, ByVal [text] As String)
        If [ControlName].InvokeRequired Then
            Dim mydelegate As New SetUIText_Delegate(AddressOf SetUIText_ThreadSafe)
            Me.Invoke(mydelegate, New Object() {[ControlName], [text]})
        Else
            [ControlName].Text = [text]
        End If
    End Sub

    Function Getparm(ByVal parmname As String, ByVal strdata As String) As String
        Dim l As Integer = Len(strdata)
        Dim spos As Integer = InStr(1, strdata, "<" & parmname & ">", CompareMethod.Text) + Len(parmname) + 5
        Dim epos As Integer = InStr(spos, strdata, "}", CompareMethod.Text)
        Dim value As String = Mid$(strdata, spos, epos - spos)
        Return value
    End Function

    Private Sub Timer1_Elapsed(ByVal sender As System.Object, ByVal e As System.Timers.ElapsedEventArgs) Handles Timer1.Elapsed

        Dim ElapsedMinutes As Long
        Dim TimeToGo As TimeSpan
        Dim TimeToDistribute As TimeSpan
        Dim TimeToGenerate As TimeSpan
        Dim TimeToSendMessages As TimeSpan
        Dim CheckFreq As Long
        Dim StartOffset As Long

        ElapsedMinutes = DateDiff("n", My.Settings.LastCheck, Now)
        CheckFreq = My.Settings.Freq
        StartOffset = My.Settings.Offset
        DistributeOffset = My.Settings.DistributeOffset
        If Not WorkOffline.Checked = True Then

            If CDate(NextCheck.Text).TimeOfDay > New TimeSpan(23, 45, 0, 0) Then
                '  Advance to Start of Next Day
                Me.LastCheck.Text = StartTime.AddDays(1)
                Me.LastDistributeCheck.Text = StartTime.AddDays(1)
                Me.NextCheck.Text = GetNextCheckTime(CDate(Me.LastCheck.Text), Offset, Freq, EndTime.AddDays(1))
                Me.NextDistribute.Text = GetNextCheckTime(CDate(Me.LastDistributeCheck.Text), DistributeOffset, DistributeFreq, EndTime.AddDays(1))
                Me.NextMessageRun.Text = GetNextCheckTime(CDate(Me.LastMessageRun.Text), DistributeOffset, DistributeFreq, EndTime.AddDays(1))
                Exit Sub
            End If


            TimeToGenerate = CDate(NextCheck.Text).TimeOfDay.Subtract(DateTime.Now.TimeOfDay)
            TimeToDistribute = CDate(NextDistribute.Text).TimeOfDay.Subtract(DateTime.Now.TimeOfDay)
            TimeToSendMessages = CDate(NextMessageRun.Text).TimeOfDay.Subtract(DateTime.Now.TimeOfDay)


            If TimeToGenerate < TimeToDistribute Then
                TimeToGo = TimeToGenerate
            Else
                TimeToGo = TimeToDistribute
            End If
            If TimeToSendMessages < TimeToGo Then
                TimeToGo = TimeToSendMessages
            End If

            If Running.Checked = False And Distributing.Checked = False And SendingMessages.Checked = False Then
                If TimeToGo.TotalMinutes < 1 Then
                    Me.StatusBar.Text = "Next Check in " & TimeToGo.Seconds.ToString & " Seconds"
                Else
                    Me.StatusBar.Text = "Next Check in " & TimeToGo.Minutes.ToString & " Minutes " & TimeToGo.Seconds.ToString & " Seconds"
                End If
            End If

            '  Start Report Generator if Overdue
            If TimeToGenerate.TotalSeconds <= 0 And Running.Checked = False Then
                Running.Checked = True
                Me.LastCheck.Text = Now.ToString
                Me.NextCheck.Text = GetNextCheckTime(StartTime, Offset, Freq, EndTime)
                Me.StatusBar.Text = "Refreshing Data..."
                refreshdata()
                '    Me.StatusBar.Text = "Starting Autoreports Run ..."
                StartReportGenerator()
            End If

            '  Start Distribution Generator 
            If TimeToDistribute.TotalSeconds <= 0 And Distributing.Checked = False Then
                Distributing.Checked = True
                Me.LastDistributeCheck.Text = Now.ToString()
                Me.NextDistribute.Text = GetNextCheckTime(StartTime, DistributeOffset, DistributeFreq, EndTime)
                '   Me.StatusBar.Text = "Starting Autoreport Distribution ..."
                StartReportDistributor()
            End If
            '  Start Messanger 
            If TimeToSendMessages.TotalSeconds <= 0 And SendingMessages.Checked = False Then
                SendingMessages.Checked = True
                Me.LastMessageRun.Text = Now.ToString()
                Me.NextMessageRun.Text = GetNextCheckTime(StartTime, DistributeOffset, DistributeFreq, EndTime)
                '   Me.StatusBar.Text = "Starting Messenger..."
                StartMessenger()
            End If

        Else
            '    Me.NextCheck.Text = "Disabled"
            '    Me.NextDistribute.Text = "Disabled"
            Me.StatusBar.Text = "Wroking Off Line or Service is Stopped"
        End If

    End Sub
    Private Sub LoadSettings()

        Dim StartTime As Date = CDate(My.Settings.StartTime)
        Dim Offset As Long = CInt(My.Settings.Offset)
        Dim Freq As Long = CInt(My.Settings.Freq)
        Dim DistributeFreq As Long = CInt(My.Settings.DistributeFreq)
        Dim DistributeOffset As Long = CInt(My.Settings.DistributeOffset)
        Dim EndTime As Date = CDate(My.Settings.EndTime)

        Me.WorkOffline.Checked = CBool(My.Settings.WorkOffline)

        Me.Running.Checked = False
        Me.RunOnTimer.Checked = CBool(My.Settings.RunOnTimer)
        Me.LastCheck.Text = My.Settings.LastCheck
        Me.NextCheck.Text = GetNextCheckTime(StartTime, Offset, Freq, EndTime)
        Me.NextDistribute.Text = GetNextCheckTime(StartTime, DistributeOffset, DistributeFreq, EndTime)
        Me.NextMessageRun.Text = GetNextCheckTime(StartTime, DistributeOffset - 1, DistributeFreq, EndTime)
    End Sub
    Public Sub refreshdata()
        Dim result As String
        Me.StatusBar.Text = "Refreshing Data"

        ' Dim ardata As New DataSet
        '  Connect to Web Service / Retrieve Schedule - If Offline, Load from local Cache

        '   Dim wstest As New AutoReportsWS3.Service1
        '   Dim ds As DataSet
        '   ds = wstest.GetAutoReportsUsers
        '   Debug.WriteLine(ds.Tables.Count.ToString)

        If Not Me.StartOffline.Checked Then
            If wsavailable() Then
                Me.Online.Checked = True
            Else
                Me.Online.Checked = False
            End If
            If Me.Online.Checked And Not WorkOffline.Checked Then
                '    Dim artest As New DataSet
                '    artest = ar.GetAutoReportsData
                '    MsgBox(artest.Tables.Count.ToString)
                dgvOverdueJobs.DataSource = ws.GetOverdueJobsList(3, DateTime.Now.ToString())
                dgvDistributionList.DataSource = ws.GetJobsToDistribute(40, DateTime.Now.AddDays(-2))
                dgvMessageList.DataSource = ws.GetOverdueMessages(10, DateTime.Now.AddDays(-1))
                dgvAccessAlerts.DataSource = ws.GetOverdueAlerts(20, DateTime.Now.AddDays(-7))

                '    ardata = ar.GetAutoReportsData
                Online.Checked = True
            Else
                Online.Checked = False
            End If
        Else
            Me.Online.Checked = False
        End If
        If Online.Checked = False Or Me.StartOffline.Checked Then
            '   Dim fsReadXml As New System.IO.FileStream("ardata.xml", System.IO.FileMode.Open)
            ' Create an XmlTextReader to read the file.
            '   Dim myXmlReader As New System.Xml.XmlTextReader(fsReadXml)
            ' Read the XML document into the DataSet.
            '   Dim artest As New DataSet
            '   artest.Clear()
            '   If System.IO.File.Exists("c:\ardata.xml") Then
            '   artest.ReadXml(myXmlReader, XmlReadMode.Auto)
            '                artest.ReadXml(
            '    MsgBox(artest.Tables.Count.ToString)
            '    artest.DataSetName.ToString()
            '  Else
            '      MsgBox("File Not Found")
            '  End If
            If "a" = "b" Then
                ardata.Clear()
                ardata.ReadXml("c:\ardata.xml", XmlReadMode.DiffGram)
                Debug.WriteLine(ardata.Tables.Count)
            Else
                '   Dim ardataLocal As New DataSet
                '        ardata.EnforceConstraints = False
                '        daApplication.Fill(ardata.APPLICATION)
                '        daReport.Fill(ardata.REPORT)
                '        daJob.Fill(ardata.JOB)
                '        Debug.WriteLine(ardata.JOB.Rows.Count)
                '        daUsers.Fill(ardata.USERS)
                ' daHistory.Fill(ardata.HISTORY)
                '        daDistributionList.Fill(ardata.DISTRIBUTION_LIST)
                '        daSetupOptions.Fill(ardata.SETUP_OPTIONS)
                '       ardata = ardataLocal.Copy
            End If
            '    MsgBox(ardata.Tables.Count.ToString)
            '  ardata.Clear()
            '  ardata.ReadXml(myXmlReader)
            '  MsgBox(myXmlReader.ReadString())
            ' Close the XmlTextReader
            '   myXmlReader.Close()
        End If

        '   If wsavailable() And StartOffline.Checked Then
        '   StartOffline.Checked = False
        '   End If
        '
        ' Setup / Refresh Display
        Try

            '  updateTreeview("all")

            DataGrid1.DataSource = ws.GetJobsList(False, False, True, True, #1/1/2010#)
            '  DataView1.RowFilter = "Active = True"
            '  DataView1.Sort = "Next_Sched,Job"
            '   DataGrid1.DataSource = ardata.Tables("SCHEDULE")
            '  dgActivityLog.DataSource = ardata.Tables("HISTORY")
            '  dgUserList2.DataSource = ardata.Tables("USERS")
            '  dgUserSubscriptions.DataSource = ardata.Tables("DISTRIBUTION_LIST")
            '  DataGrid1.DataSource = ardata.Tables("REPORT")
            '  dgASNList.DataSource = ardata.Tables("NONRECURRING_JOBS")
            Me.Refresh()
            result = "Updated"
        Catch ex As Exception
            '     StatusBar.Text = ex.ToString
            result = "Not Updated"
        End Try
    End Sub
    Private Function wsavailable() As Boolean
        Try

            Debug.WriteLine(ws.HelloWorld(DateTime.Now).ToString)
            If Len(ws.HelloWorld(DateTime.Now.ToLocalTime)) > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function
    Private Sub RunOverdueReports(ByVal Order As String)
        '  Setup / Start Timer
        Dim CheckStart As Date
        Dim ws As New AutoReportsWCFService.ServiceClient
        Dim al() As AppsList
        Dim jl() As JobList
        Dim reportoutcome As Boolean
        Dim uploadoutcome As Boolean
        Dim jobcompletion As DateTime
        Dim archcontainer As String

        Dim strtempdirectory As String
        Dim strtempfilename As String
        Dim strtempextension As String
        Dim strtempfilepath As String
        Dim archfilename As String
        Dim strOutputFormat As String
        Dim ArchiveID As Integer
        Try
            CheckStart = Now
            Timer1.Enabled = False
            al = ws.GetAppsList()

            For Each app As AppsList In al
                Me.StatusBar.Text = "Checking App " + app.DESCRIPTION + "...."
                jl = ws.GetOverdueJobsList(app.APPID, DateTime.Now)
                If jl.Length > 0 Then
                    '  Open App
                    oAccess = OpenAccessApplication(app.STARTUP_DIRECTORY, app.STARTUP_FILENAME, app.STARTUP_OPTIONS, app.MDW_DIRECTORY + app.MDW_FILENAME, app.USERNAME, app.PASSWORD)

                    For Each OverdueJob As JobList In jl
                        Me.StatusBar.Text = "Running Job " + OverdueJob.DESCRIPTION + " Scheduled for " + OverdueJob.NEXT_SCHED.ToString()
                        Dim applicationpath As String = app.STARTUP_DIRECTORY + app.STARTUP_FILENAME

                        strtempdirectory = "C:\HTS\REPORTS\"
                        strtempfilename = "AR" + Format(OverdueJob.JOBID, "000000") + ".OUT"
                        strtempfilepath = strtempdirectory & strtempfilename

                        reportoutcome = GenerateOutput(OverdueJob.JOBID, OverdueJob.REPORT_NAME, OverdueJob.CRITERIA, strtempdirectory, strtempfilename, OverdueJob.OUTPUT_FORMAT)


                        strOutputFormat = OverdueJob.OUTPUT_FORMAT

                        Dim attachfilename As String = Format(OverdueJob.JOBID, "000000") + " - " + OverdueJob.DESCRIPTION + "." + OverdueJob.OUTPUT_FORMAT
                        Dim attachmentfilepath As String = "c:\HTS\REPORTS\" & attachfilename

                        jobcompletion = Now.ToLocalTime()
                        archcontainer = OverdueJob.CONTAINER
                        archfilename = Format(jobcompletion, "yyyyMMddHHmm") & "." & strOutputFormat

                        If reportoutcome = True Then
                            ' Save local copy for email transmission / server upload
                            If System.IO.File.Exists(attachmentfilepath) Then
                                System.IO.File.Delete(attachmentfilepath)
                            End If
                            Thread.Sleep(1000)
                            System.IO.File.Copy(strtempfilepath, attachmentfilepath)
                            Thread.Sleep(1000)
                            '  Updoad to Azure
                            If Online.Checked Then
                                Dim shortdesc As String = "(" & OverdueJob.JOBID & ") " + OverdueJob.DESCRIPTION + " @ " + jobcompletion.ToShortTimeString()
                                StatusBar.Text = "Uploading... " + shortdesc + " to Azure"
                                uploadoutcome = UploadReportToAzure(strtempdirectory, strtempfilename, archcontainer, archfilename)
                                ArchiveID = ws.LogArchive(OverdueJob.TYPE, archcontainer, archfilename, OverdueJob.OUTPUT_FORMAT, OverdueJob.JOBID, jobcompletion, OverdueJob.NEXT_SCHED, "TBD", shortdesc, "")
                            End If
                        End If
                    Next
                    '  Close App
                    CloseAccessApplication(oAccess)
                Else
                    StatusBar.Text = "No overdue Reports for App " = app.DESCRIPTION
                End If
            Next

        Catch ex As Exception
            Me.StatusBar.Text = "Error: " + ex.Message.ToString()
        Finally
            Timer1.Enabled = True
        End Try

    End Sub

    Public Function UploadReportToAzure(ByVal LocalDirectory As String, ByVal LocalFileName As String, ByVal AzureContainerName As String, ByVal AzureFileName As String) As Boolean
        Try
            ' Setup Variables for Cloud Storage Objects
            Dim cloudStorageAccount As CloudStorageAccount
            Dim blobClient As CloudBlobClient
            Dim blobContainer As CloudBlobContainer
            Dim containerPermissions As BlobContainerPermissions
            Dim blob As CloudBlob
            Dim blob2 As CloudBlockBlob


       '     Dim cs As String
       '     cs = "DefaultEndpointsProtocol=https;AccountName=htsazure;AccountKey=pbhjjJe9/tHpkcNq21UBlRWi+MYTr00HtUUVkqOe4Z+mqFzfF57IcF8PaD/FaWkE6Y0cOnntjd0mmJmwDDwyKg=="
       '     cs = "DefaultEndpointsProtocol=http;AccountName=sensible;AccountKey=jTiCHXRtXPdEAsnbEhUXAAdNWTwVrR6KeZvyfT6qYusAfyqDKZn8R2Pqy3Os3yYbSj8Zj7zaW5KMohyTr3dMkg=="
       '     cloudStorageAccount = cloudStorageAccount.Parse(cs)
       '     blobClient = cloudStorageAccount.CreateCloudBlobClient()
       '     blobContainer = blobClient.GetContainerReference("0010")
       '     blob2 = blobContainer.GetBlockBlobReference("testfile")
       '     Using fileStream = System.IO.File.OpenRead("C:\uploadfolder\output.csv")
       '         blob2.UploadFromStream(fileStream)
       '     End Using


            ' Use the emulated storage account or use the real account
            ' cloudStoargeAccount = CloudStorageAccount.DevelopmentStorageAccount

            ' Use the Windows Azure cloud storage account
         '   cloudStoargeAccount = CloudStorageAccount.Parse("DefaultEndpointsProtocol=http;AccountName=sensible;AccountKey=jTiCHXRtXPdEAsnbEhUXAAdNWTwVrR6KeZvyfT6qYusAfyqDKZn8R2Pqy3Os3yYbSj8Zj7zaW5KMohyTr3dMkg==")
         '   cloudStorageAccount = cloudStorageAccount.Parse("DefaultEndpointsProtocol=https;AccountName=htsazure;AccountKey=pbhjjJe9/tHpkcNq21UBlRWi+MYTr00HtUUVkqOe4Z+mqFzfF57IcF8PaD/FaWkE6Y0cOnntjd0mmJmwDDwyKg==")
             cloudStorageAccount = cloudStorageAccount.Parse("DefaultEndpointsProtocol=https;AccountName=htsazure;AccountKey=SYowHJqWJtHVau1amgFddZWyCKutTBWTRUpv5al+nlvr4sDDLBMdJXIbcFhlRJ4HS/YAtmWh5eoNTP8OBOWR2Q==")
            '  clodStorageAccount = Storage.CloudStorageAccount.Parse(Configuration

            ' cloudStoargeAccount = CloudStorageAccount.Parse("DefaultEndpointsProtocol=http;AccountName=htsazure;AccountKey=aY4Qq3xbXKXDoh7YfZmeYIdBQ9qKGu5mrAZ/PvGygeBnbDLl0+zcx1J8G3HdWALWyTqVuf7er35+P7V9F0xy8Q==;EndpointSuffix=core.windows.net")


            ' Create the blob client which priveds autheneticated access to the Blob service    
            blobClient = cloudStorageAccount.CreateCloudBlobClient()

            ' Get the container reference
            blobContainer = blobClient.GetContainerReference(AzureContainerName)

            ' Create the container if it does not exist
            blobContainer.CreateIfNotExist()

            ' Set Permissions on the container
            containerPermissions = New BlobContainerPermissions()

            ' This example sets the container to have public blobs
            containerPermissions.PublicAccess = BlobContainerPublicAccessType.Blob
            blobContainer.SetPermissions(containerPermissions)

            ' Get a reference to the blob
            blob = blobContainer.GetBlobReference(AzureFileName)

            ' Upload a file from the local system to the blob
            Debug.WriteLine("Starting File Upload to Azure...")
            blob.UploadFile(LocalDirectory + LocalFileName)
            Debug.WriteLine("File Upload Completed to Azure ..." + blob.Uri.ToString())

            Return True
        Catch ex As Exception
            Debug.WriteLine("Storage client error encountered: " + ex.Message.ToString())
            Return False
        End Try

    End Function

    Public Function SendEmailX(ByVal distlist As UserList(), ByVal from As String, ByVal Subject As String, ByVal BodyText As String, ByVal AttachFile As String) As Boolean

        Dim sattach As String = AttachFile
        '        Dim FromUserName As String = System.Configuration.ConfigurationManager.AppSettings("FromUserName")
        Dim FromUserName As String = My.Settings.FromUserName
        '        Dim SMTPServerName As String = System.Configuration.ConfigurationManager.AppSettings("SMTPServerName")
        Dim SMTPServerName As String = My.Settings.SMTPServerName
        '   Dim ccDistList As String = System.Configuration.ConfigurationManager.AppSettings("ccDistList")
        Dim ccDistList As String = My.Settings.ccDistList
        '        Dim SMTPUser As String = System.Configuration.ConfigurationManager.AppSettings("SMTPUser")
        Dim SMTPUser As String = My.Settings.SMTPUser
        '       Dim SMTPPassword As String = System.Configuration.ConfigurationManager.AppSettings("SMTPPassword")
        Dim SMTPPassword As String = My.Settings.SMTPPassword
        '  SMTPServerName = "HTSSERVER.htsmi.local"
        '  SMTPUser = "edolikian"
        ' SMTPPassword = "Sens1ble"
        ' ccDistList = "autoreports@ssitroy.com,fshepard@htsmi.com,jwhaley@htsmi.com"
        ' FromUserName = "autoreports@ssitroy.com"
        Dim email As New MailMessage()
        Try
            If Len(sattach) > 0 Then
                Dim myAttachment As Mail.Attachment = New Mail.Attachment(sattach)
                email.Attachments.Add(myAttachment)
            End If
            Thread.Sleep(100)
            If "a" = "a" Then
                For Each person As UserList In distlist
                    email.To.Add(New MailAddress(person.EMAIL, person.FULLNAME))
                Next
            Else
                email.To.Add(New MailAddress("autoreports@ssitroy.com", "Ed Dolikian"))
            End If
            email.From = New MailAddress(FromUserName, "HTS Online Reports")
            email.Subject = Subject
            email.Body = BodyText
            email.IsBodyHtml = True
            

            '  Need to split ccDistList by ";"
            Dim ccList As String() = ccDistList.Split(";")
            For Each item As String In ccList
                email.CC.Add(item)
            Next
            '  email.CC.Add(ccDistList)
          '  SMTPServerName = "smtp.office365.com"
          '  SMTPUser = "autoreports@htsmi.com"
          '  SMTPPassword = "Getmeintoit5%"



            Dim mailClient As New Mail.SmtpClient(SMTPServerName)
            mailClient.Credentials = New System.Net.NetworkCredential(SMTPUser, SMTPPassword)
            mailClient.Port = 587
            mailClient.Send(email)
            Return True
            SetUIText_ThreadSafe(Me.StatusBar, "Email Sent")
            Return True
        Catch ex As Exception
            Debug.WriteLine(ex.Message.ToString)
            SetUIText_ThreadSafe(Me.StatusBar, ex.Message.ToString)
            Return False
            ' This converted to True to mark as sent during sbsserver changeover on 2/11/11 / need to change back once server is back up
            '            Return False
        Finally
            Thread.Sleep(100)
        End Try

    End Function
    
    Public Function SendEmailToGroup(ByVal distlist As string, ByVal from As String, ByVal Subject As String, ByVal BodyText As String, ByVal AttachFile As String) As Boolean

        Dim sattach As String = AttachFile
        Dim FromUserName As String = My.Settings.FromUserName
        Dim SMTPServerName As String = My.Settings.SMTPServerName
        Dim ccDistList As String = My.Settings.ccDistList
        Dim SMTPUser As String = My.Settings.SMTPUser
        Dim SMTPPassword As String = My.Settings.SMTPPassword
        Dim email As New MailMessage()
        Try
            If Len(sattach) > 0 Then
                Dim myAttachment As Mail.Attachment = New Mail.Attachment(sattach)
                email.Attachments.Add(myAttachment)
            End If
            Thread.Sleep(100)
            Dim SendTo() As String
            Dim N As Integer
            SendTo = distlist.Split(";")
            For Each Item As String In SendTo
                email.To.Add(item)
            Next
            email.From = New MailAddress(FromUserName, "HTS Online Reports")
            email.Subject = Subject
            email.Body = BodyText
            email.IsBodyHtml = True

            '  Need to split ccDistList by ";"
            Dim ccList As String() = ccDistList.Split(";")
            For Each item As String In ccList
                email.CC.Add(item)
            Next
          
            Dim mailClient As New Mail.SmtpClient(SMTPServerName)
            mailClient.Credentials = New System.Net.NetworkCredential(SMTPUser, SMTPPassword)
            mailClient.Port = 587
            mailClient.Send(email)
            Return True
            SetUIText_ThreadSafe(Me.StatusBar, "Email Sent")
            Return True
        Catch ex As Exception
            Debug.WriteLine(ex.Message.ToString)
            SetUIText_ThreadSafe(Me.StatusBar, ex.Message.ToString)
            Return False
            ' This converted to True to mark as sent during sbsserver changeover on 2/11/11 / need to change back once server is back up
            '            Return False
        Finally
            Thread.Sleep(100)
        End Try

    End Function
    

    Private Function DistributeAutoReports(ByVal MaxCount As Long, ByVal SinceDate As DateTime) As Long


        ' Retrieve List of Jobs to Distribute
        Dim ws As New AutoReportsWCFService.ServiceClient
        Dim al() As ArchiveList
        Dim dl() As UserList
        Dim person As UserList
        Dim NumberSent As Long = 0
        Dim NumberToSend As Long = 0

        '  Retrieve List of Jobs to Distribute
        al = ws.GetJobsToDistribute(MaxCount, SinceDate.AddDays(-2))
        NumberToSend = al.Length()
        '  Process Each Archive
        For Each row As ArchiveList In al
            NumberSent += 1

            ' Determine Distribution List
            If row.JOBTYPE = "ASN" Then
                dl = ws.GetASNDistributionList(row.JOBREF)
            Else
                dl = ws.GetJobDistributionList(row.JOBREF)
            End If

            ' Prepare Message to Send to Distribution List
            Dim specialmessage As String

            Dim subject As String = row.SUBJECT + " @ " + row.CREATED.ToShortDateString + " " + row.CREATED.ToShortTimeString
            Dim strfrom As String
            Dim Note As String = "DO NOT REPLY TO SENDER (To Reply, Click on a Recipient's email or All Recipients below)"
            '    Dim documenturi As String = "http://sensible.blob.core.windows.net/" + row.CONTAINER + "/" + row.FILENAME
            Dim documenturi As String = "http://htszaure.blob.core.windows.net/" + row.CONTAINER + "/" + row.FILENAME
            Dim websiteuri As String = "http://75.151.4.117/OnlineReports"
            Dim localfile As String = "C:\HTS\REPORTS\" + "AR" + Format(row.JOBREF, "000000")
            Dim attachfile As String = "C:\HTS\REPORTS\" + "AR" + Format(row.JOBREF, "000000") + "." + row.JOBFORMAT


            '  Create Attachment w/ useful filename
            If row.JOBFORMAT = "PDF" Then
                localfile = localfile + ".OUT"

                ' Copy localfile to create file to send
                Thread.Sleep(100)

                If System.IO.File.Exists(attachfile) Then
                    System.IO.File.Delete(attachfile)
                End If
                If localfile <> attachfile Then
                    System.IO.File.Copy(localfile, attachfile)
                End If

                Thread.Sleep(100)

            Else
                localfile = localfile + ".XLS"
            End If

            '  strfrom = "autoreports@ssitroy.com"
            strfrom = "autoreports@htsmi.com"

            specialmessage = "<p>IMPORTANT NOTE</p>"
            If Note.Length > 0 Then
                specialmessage = specialmessage & Note & Chr(10) & Chr(13)
            End If

            specialmessage = specialmessage & "<p><a href=" & documenturi & ">Click Here to View Report</a></p>" & Chr(10) & Chr(13)

            specialmessage = specialmessage & "<p>You are receiving this automated email from Heat Treating Services.  If you would like to be removed from distribution or would like someone else to be added, please contact Franklin Shepard or simply reply to this email.</p>" & Chr(10) & Chr(13)
            specialmessage = specialmessage & "<p>To ensure proper delivery, please be sure to add fshepard@htsmi.com to your trusted senders list.</p>" & Chr(10) & Chr(13)
            specialmessage = specialmessage & "<p>For any production or scheduling questions, please contact your plant representative or call (248) 858-2230.</p>" & Chr(10) & Chr(13)
            specialmessage = specialmessage & "<p>Thank you</p>"

            ' specialmessage = specialmessage + "<p>Click here to Login to <a href=" + websiteuri + "> HTS Online</a></p>" + Chr(10) & Chr(13)

            ' Update Activity Log / Populate Recipients
            Dim distlist As String = ""

            specialmessage = specialmessage & "<p>Current Distribution List:</p>"
            specialmessage = specialmessage & "<ul>"
            For Each person In dl
                If distlist <> "" Then
                    distlist = distlist + ","
                End If
                distlist = distlist + person.EMAIL
                specialmessage = specialmessage & "<li>" & person.FULLNAME & " (" & person.EMAIL & ")"
                ws.LogActivity(row.JOBTYPE, row.ARCHIVEID, person.USERID, row.JOBREF, "Email Sent at " + Now.ToShortTimeString(), DateTime.Now)
            Next
            specialmessage = specialmessage & "</ul>"

            SetUIText_ThreadSafe(Me.StatusBar, "Emailing..." + NumberSent.ToString + " of " + NumberToSend.ToString + " " + row.JOBREF.ToString + " " + row.SUBJECT & " to " + distlist)

            specialmessage = specialmessage & "<p>Click here to send a message to <a href=mailto:" & distlist & "?subject=Reply%20Message%20to%20AutoReport>" & "All Recipients</a></p>" & Chr(10) & Chr(13)

            '     specialmessage = specialmessage & "<p>Current Distribution List => " & distlist & "</p>" & Chr(10) & Chr(13)

            specialmessage = specialmessage & "<p>Click here to send a message to <a href=mailto:autoreports@ssitroy.com;fshepard@htsmi.com" + "?subject=Message%20to%20Admins%20Message%20to%20AutoReport>" + "System Administrators</a></p>" & Chr(10) & Chr(13)

            specialmessage = specialmessage & "<p>Message Sent from HTSAUTO at " & DateTime.Now.ToString & " / Ref: " & row.ARCHIVEID & "</p>" & Chr(10) & Chr(13)
            ' Send Email
            Dim result As Boolean = SendEmailX(dl, strfrom, subject, specialmessage, attachfile)


            ' Mark Archive as Sent
            ws.MarkArchiveAsSent(row.ARCHIVEID, distlist)
            If row.JOBTYPE = "ASN" Then
                ws.MarkASNAsSent(row.JOBREF, DateTime.Now, row.CONTAINER, row.FILENAME, distlist)
            End If

        Next
        Return NumberSent
    End Function


    Private Sub updateTreeview(ByVal View As String)
        '   TreeView2.BeginUpdate()

        TreeView1.Nodes.Clear()

        TreeView1.Nodes.Add(New TreeNode("All Applications"))
        Dim appnode As TreeNode = TreeView1.Nodes(0)
        Dim dr As DataRow
        Dim dt As DataTable
        dt = ardata.Tables("APPLICATION")
        Dim i As Long
        For i = 0 To dt.Rows.Count - 1
            appnode.Nodes.Insert(i, New TreeNode(dt.Rows(i).Item("Description").ToString))
        Next i

        TreeView1.Nodes.Add(New TreeNode("All Users"))
        Dim usernode As TreeNode = TreeView1.Nodes(1)
        dt = ardata.Tables("USERS")
        For i = 0 To dt.Rows.Count - 1
            usernode.Nodes.Insert(i, New TreeNode(dt.Rows(i).Item("Contact").ToString))
        Next i

        TreeView1.Nodes.Add(New TreeNode("Jobs"))
        Dim rptnode As TreeNode = TreeView1.Nodes(2)
        dt = ardata.Tables("JOB")
        For i = 0 To dt.Rows.Count - 1
            rptnode.Nodes.Insert(i, New TreeNode(dt.Rows(i).Item("JOB").ToString))
        Next i

        TreeView1.Show()
        '   End If
    End Sub
    'Public Function SendEmailMessagex(ByVal distlist As String, ByVal from As String, ByVal Subject As String, ByVal BodyText As String, ByVal AttachFile As String) As Boolean

    '    Dim email As New MailMessage
    '    Dim smtp As SmtpMail
    '    Dim sattach As String = AttachFile
    '    Dim FromUserName As String = ConfigurationSettings.AppSettings("FromUserName")
    '    Dim SMTPServerName As String = ConfigurationSettings.AppSettings("SMTPServerName")
    '    Dim ccDistList As String = ConfigurationSettings.AppSettings("ccDistList")

    '    Try
    '        If Len(sattach) > 0 Then
    '            Dim myAttachment As MailAttachment = New MailAttachment(sattach)
    '            email.Attachments.Add(myAttachment)
    '        End If
    '        Thread.Sleep(1000)



    '        ' Comment out the Following after it works
    '        '     distlist = "edolikian@htsmi.com"
    '        '     from = "fshepard@htsmi.com"
    '        '  Subject = "Test Report"
    '        '  ccDistList = ""
    '        '    SMTPServerName = "HTSSERVER.htsmi.local"

    '        email.To = distlist
    '        email.From = from
    '        '   email.From = "edolikian@htsmi.com"
    '        email.Subject = Subject
    '        email.Body = BodyText
    '        email.BodyFormat = System.Web.Mail.MailFormat.Text
    '        email.Cc = ccDistList
    '        Dim mailClient As SmtpMail
    '        mailClient.SmtpServer = SMTPServerName
    '        ' mailClient.Send("edolikian@htsmi.com", "edolikian@ssitroy.com", "subject", "message")
    '        mailClient.Send(email)


    '        '   email.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate", "1")    'basic authentication
    '        '   email.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusername", "edolikian@htsmi.com") 'set your username here
    '        '   email.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendpassword", "Sens1ble") 'set your password here
    '        '   email.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpserverport", "9925") ' put your port number here
    '        '   smtp.SmtpServer.Insert(0, SMTPServerName) 'your real server goes here
    '        '  smtp.SmtpServer = SMTPServerName

    '        '   smtp.SmtpServer.Insert(0, "mail.comcast.net")

    '        '  smtp.Send(email)
    '        Return True





    '        '  smtp.SmtpServer.Insert(0, SMTPServerName)
    '        '  smtp.Send(email)


    '        '       System.Web.Mail.SmtpMail.SmtpServer = SMTPServerName

    '        '        System.Web.Mail.SmtpMail.Send(EMAIL)
    '        '     System.Web.Mail.SmtpMail.SmtpServer = "mail.comcast.net"
    '        '  System.Web.Mail.SmtpMail.SmtpServer = "mail.sensiblesolutionsinc.net"

    '        '    EMAIL.Cc = "edolikian@ssitroy.com"
    '        '  EMAIL.From = "edolikian@ssitroy.com"

    '        '          System.Web.Mail.SmtpMail.Send(EMAIL)
    '        StatusBar.Text = "Email Sent"
    '        Return True
    '    Catch ex As Exception
    '        Debug.WriteLine(ex.Message.ToString)
    '        StatusBar.Text = ex.Message.ToString
    '        Return False
    '        ' This converted to True to mark as sent during sbsserver changeover on 2/11/11 / need to change back once server is back up
    '        '            Return False
    '    Finally
    '        Thread.Sleep(2000)
    '    End Try

    'End Function
    'Public Function SendEmailMessageZ()
    '    '  Dim email As New System.Web.Mail.MailMessage
    '    Dim email As New System.Net.Mail.MailMessage("edolikian@htsmail.com", "edolikian@htsmi.com", "Subject", "Body")

    '    '   email.To = "edolikian@ssitroy.com"
    '    '   email.From = "edolikian@htsmi.com"
    '    email.Body = "MessageText"
    '    email.Subject = "SubjectText"
    '    '     email.BodyFormat = System.Net.Web.Mail.MailFormat.Text
    '    '   System.Net.Mail.s.SmtpMail.SmtpServer.Insert(0, "HTSSERVER\htsmi.local")
    '    '   System.Net.Mail.MailMessa.S.Web.Mail.SmtpMail.Send(email)
    '    SmtpMail.Send(email)
    'End Function

    'Public Function SendEmailMessagey(ByVal distlist As String, ByVal from As String, ByVal Subject As String, ByVal BodyText As String, ByVal AttachFile As String) As Boolean

    '    Dim email As New MailMessage
    '    Dim smtp As System.Net.Mail.MailMessage
    '    smtp.SmtpServer.Insert(0, "HTSSERVER.htsmi.local")
    '    ' smtp.SmtpServer.Insert(0, "192.168.1.200:9925")
    '    Dim sattach As String = AttachFile
    '    Dim FromUserName As String = ConfigurationSettings.AppSettings("FromUserName")
    '    Dim SMTPServerName As String = ConfigurationSettings.AppSettings("SMTPServerName")
    '    Dim ccDistList As String = ConfigurationSettings.AppSettings("ccDistList")


    '    '  this is the new test code

    '    Try
    '        If Len(sattach) > 0 Then
    '            Dim myAttachment As MailAttachment = New MailAttachment(sattach)
    '            email.Attachments.Add(myAttachment)
    '        End If
    '        Thread.Sleep(1000)

    '        email.To = distlist
    '        email.From = from
    '        email.Subject = Subject
    '        email.Body = BodyText + smtp.SmtpServer.ToString

    '        email.BodyFormat = System.Web.Mail.MailFormat.Text
    '        email.Cc = ccDistList

    '        smtp.SmtpServer.Insert(0, SMTPServerName)
    '        smtp.Send(email)


    '        '       System.Web.Mail.SmtpMail.SmtpServer = SMTPServerName

    '        '        System.Web.Mail.SmtpMail.Send(EMAIL)
    '        '     System.Web.Mail.SmtpMail.SmtpServer = "mail.comcast.net"
    '        '  System.Web.Mail.SmtpMail.SmtpServer = "mail.sensiblesolutionsinc.net"

    '        '    EMAIL.Cc = "edolikian@ssitroy.com"
    '        '  EMAIL.From = "edolikian@ssitroy.com"

    '        '          System.Web.Mail.SmtpMail.Send(EMAIL)
    '        StatusBar.Text = "Email Sent"
    '        Return True
    '    Catch ex As Exception
    '        Debug.WriteLine(ex.Message.ToString)
    '        StatusBar.Text = ex.Message.ToString
    '        Return False
    '        ' This converted to True to mark as sent during sbsserver changeover on 2/11/11 / need to change back once server is back up
    '        '            Return False
    '    Finally
    '        Thread.Sleep(2000)
    '    End Try

    'End Function
    Public Function GenerateReport(ByVal ApplicationPath As String, ByVal ReportName As String, ByVal Criteria As String, ByVal User As String, ByVal Password As String, ByVal stroutputdirectory As String, ByVal stroutputfilename As String, ByVal stroutputformat As String) As Boolean
        Dim outcome As Boolean
        Try
            Call Print_Report_Security(ApplicationPath, ReportName, Criteria, User, Password, stroutputdirectory, stroutputfilename, stroutputformat)
            outcome = True
        Catch ex As Exception
            StatusBar.Text = ex.Message & " Detected in Catch Block of Generate Report Routine"
            outcome = False
        Finally
            GenerateReport = outcome
        End Try
    End Function
    Private Function GetNextSchedule(ByVal freq As String, ByVal qty As Long, ByVal start_date As Date, ByVal start_time As Date, ByVal end_time As Date, ByVal dayoffset As Long) As Date
        ' This function returns the next scheduled runtime based on the current date and time
        Dim lastsched As String = ""
        Dim nextsched As String
        Dim cutoff As String
        Try
            Select Case freq
                Case Is = "h", "n", "N", "H"
                    ' This is a multiple times per day type
                    lastsched = Today.ToShortDateString + " " + start_time.ToShortTimeString
                    nextsched = Today.ToShortDateString + " " + start_time.ToShortTimeString
                    cutoff = Today.ToShortDateString + " " + "11:00 pm"
                    cutoff = Today.ToShortDateString + " " + end_time.ToShortTimeString
                    Do Until DateDiff("n", nextsched, Now) < 0
                        lastsched = nextsched
                        nextsched = DateAdd(freq, qty, lastsched)
                    Loop
                    If CDate(nextsched) >= CDate(cutoff) Then
                        ' Next Scheduled is after Ending Time / Set to first time of next day
                        nextsched = DateAdd("d", 1, CDate(DateValue(lastsched) + " " + start_time))
                    End If

                Case Else
                    ' This is a once per day or schedule type
                    lastsched = DateValue(start_date) & " " & start_time
                    nextsched = lastsched
                    Do Until DateDiff("n", DateAdd("d", dayoffset, nextsched), Now) < 0
                        nextsched = DateAdd(freq, qty, lastsched)
                        lastsched = nextsched
                    Loop

                    nextsched = DateAdd("d", dayoffset, lastsched)

            End Select
            'Return nextsched
        Catch ex As Exception
            Me.StatusBar.Text = ex.Message
            nextsched = lastsched
        Finally
            GetNextSchedule = nextsched
        End Try

    End Function
    Public Function ResetNextScheduleDate() As Boolean
        Try
            Dim dtSchedule As DataTable = ardata.Tables("SCHEDULE")
            Dim reccnt As Long = dtSchedule.Rows.Count
            Dim i As Long
            Dim dr As DataRow
            For i = 0 To reccnt - 1
                dr = dtSchedule.Rows(i)
                dr("Next_Sched") = GetNextSchedule(dr("Freq"), dr("qty"), dr("Start_Date"), dr("start_time"), dr("end_time"), dr("start_offset"))
                Dim tod = TimeValue(Now)
            Next i
            Return True
        Catch ex As Exception
            Return False
        End Try

        '   daSchedule.Update(DataSet2.SCHEDULE)
    End Function
    Private Sub DataGrid1_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs)


    End Sub
    Public Function ShellGetDB(ByVal sDBPath As String,
       Optional ByVal sCmdLine As String = vbNullString,
       Optional ByVal enuWindowStyle As Microsoft.VisualBasic.AppWinStyle _
           = AppWinStyle.MinimizedFocus,
       Optional ByVal iSleepTime As Integer = 1000) As Microsoft.Office.Interop.Access.Application

        Dim oAccess As Access.Application = New Access.Application
        Dim sAccPath As String 'path to msaccess.exe

        ' Obtain the current process name
        Dim myprocess As Process
        myprocess = Process.GetCurrentProcess
        '  Dim myprocessname As String = myprocess.ProcessName
        ' Obtain the path to msaccess.exe:
        Try
            '   MessageBox.Show("Before GetOfficeAppPath")
            '   sAccPath = GetOfficeAppPath("Access.Application", "msaccess.exe")
            sAccPath = "C:\Program Files (x86)\Microsoft Office\Office14\MSACCESS.EXE"

            If sAccPath = "" Then
                MsgBox("Can't determine path to msaccess.exe",
                    MsgBoxStyle.MsgBoxSetForeground)
                Return Nothing
            End If

            ' Make sure specified database (sDBPath) exists:
            If Not System.IO.File.Exists(sDBPath) Then
                MsgBox("Can't find the file '" & sDBPath & "'",
                    MsgBoxStyle.MsgBoxSetForeground)
                Return Nothing
            End If

            ' Start a new instance of Access using sDBPath and sCmdLine:
            If sCmdLine = vbNullString Then
                sCmdLine = Chr(34) & sDBPath & Chr(34)
            Else
                sCmdLine = Chr(34) & sDBPath & Chr(34) & " " & sCmdLine
            End If

            '    MessageBox.Show("Before Shell Command sAccpath=" & sAccPath & " scmdline = " & sCmdLine)

            Shell(PathName:=sAccPath & " " & sCmdLine & " " & sDBPath & " /user autoreports /pwd 4438", Style:=enuWindowStyle)


            '       Shell(PathName:=sAccPath & " " & sCmdLine, Style:=enuWindowStyle)

            ' Pause to allow database to open:
            '    iSleepTime = 6000
            iSleepTime = 2000
            System.Threading.Thread.Sleep(iSleepTime)
            '    AppActivate(Title:=myprocessname)
            ' Obtain Application object of the instance of Access
            ' that has the database open:
            '  MessageBox.Show("Before GetObject Command")
            oAccess = GetObject(sDBPath)
            '   oAccess = GetObject(, "Access Application")
            Return oAccess
        Catch ex As Exception
            '   MessageBox.Show("In ShellGetDB Execption Block" & ex.Message)
            oAccess.Quit(Option:=Access.AcQuitOption.acQuitSaveNone)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oAccess)
            oAccess = Nothing

            Return Nothing

            StatusBar.Text = ex.Message
        End Try
    End Function

    Public Function ShellGetApp(Optional ByVal sCmdLine As String = vbNullString,
        Optional ByVal enuWindowStyle As Microsoft.VisualBasic.AppWinStyle _
            = AppWinStyle.MinimizedFocus) As Access.Application

        'Launches a new instance of Access using the Shell function
        'then returns the Application object via calling:
        'GetObject(,"Access.Application"). If an instance of
        'Access is already running before calling this procedure,
        'the function may return the Application object of a
        'previously running instance of Access. If this is not
        'desired, then make sure Access is not running before
        'calling this function, or use the ShellGetDB()
        'function instead. Approach based on Q308409.
        '
        'Examples:
        'Dim oAccess As Access.Application
        'oAccess = ShellGetApp()
        '
        '-or-
        '
        'Dim oAccess As Access.Application
        'Dim sUser As String
        'Dim sPwd As String
        'sUser = "user_name"
        'sPwd = "my_password"
        'oAccess = ShellGetApp("/user " & sUser & "/pwd " & sPwd)

        ' Enable an error handler for this procedure:
        On Error GoTo ErrorHandler

        Dim oAccess As Access.Application
        Dim sAccPath As String 'path to msaccess.exe
        Dim iSection As Integer = 0
        Dim iTries As Integer = 0

        ' Obtain the path to msaccess.exe:
        sAccPath = GetOfficeAppPath("Access.Application", "msaccess.exe")
        If sAccPath = "" Then
            MsgBox("Can't determine path to msaccess.exe",
                MsgBoxStyle.MsgBoxSetForeground)
            Return Nothing
        End If

        ' Start a new instance of Access using sCmdLine:
        If sCmdLine = vbNullString Then
            sCmdLine = sAccPath
        Else
            sCmdLine = sAccPath & " " & sCmdLine
        End If
        Shell(PathName:=sCmdLine, Style:=enuWindowStyle)
        'Note: It is advised that the Style argument of the Shell
        'function be used to give focus to Access.

        ' Move focus back to this form. This ensures that Access
        ' registers itself in the ROT, allowing GetObject to find it:
        AppActivate(Title:=Me.Text)

        ' Attempt to use GetObject to reference a running
        ' instance of Access:
        iSection = 1 'attempting GetObject...
        oAccess = GetObject(, "Access.Application")
        iSection = 0 'resume normal error handling

        Return oAccess
ErrorCleanup:
        ' Try to quit Access due to an unexpected error:
        On Error Resume Next
        oAccess.Quit(Option:=Access.AcQuitOption.acQuitSaveNone)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oAccess)
        oAccess = Nothing

        Return Nothing
ErrorHandler:
        If iSection = 1 Then 'GetObject may have failed because the
            'Shell function is asynchronous; enough time has not elapsed
            'for GetObject to find the running Office application. Wait
            '1/2 seconds and retry the GetObject. If you try 20 times
            'and GetObject still fails, assume some other reason
            'for GetObject failing and exit the procedure.
            iTries = iTries + 1
            If iTries < 20 Then
                System.Threading.Thread.Sleep(500) 'wait 1/2 seconds
                AppActivate(Title:=Me.Text)
                Resume 'resume code at the GetObject line
            Else
                MsgBox("GetObject failed. Process ended.",
                    MsgBoxStyle.MsgBoxSetForeground)
            End If
        Else 'iSection = 0 so use normal error handling:
            MsgBox(Err.Number & ": " & Err.Description,
                MsgBoxStyle.MsgBoxSetForeground, "Error Handler")
        End If
        Resume ErrorCleanup
    End Function

    Public Function GetOfficeAppPath(ByVal sProgId As String, ByVal sEXE As String) As String
        'Returns path of the Office application. e.g.
        'GetOfficeAppPath("Access.Application", "msaccess.exe") returns
        'full path to Access. Approach based on Q240794.
        'Returns empty string if path not found in registry.

        ' Enable an error handler for this procedure:

        Dim oReg As Microsoft.Win32.RegistryKey =
            Microsoft.Win32.Registry.LocalMachine
        Dim oKey As Microsoft.Win32.RegistryKey
        Dim sCLSID As String
        Dim sPath As String
        Dim iPos As Integer

        ' First, get the clsid from the progid from the registry key
        ' HKEY_LOCAL_MACHINE\Software\Classes\<PROGID>\CLSID:
        Try
            oKey = oReg.OpenSubKey("Software\Classes\" & sProgId & "\CLSID")

            sCLSID = oKey.GetValue("")
            oKey.Close()

            ' Now that we have the CLSID, locate the server path at
            ' HKEY_LOCAL_MACHINE\Software\Classes\CLSID\ 
            ' {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxx}\LocalServer32:
            oKey = oReg.OpenSubKey("Software\Classes\CLSID\" & sCLSID & "\LocalServer32")
            sPath = oKey.GetValue("")
            oKey.Close()

            ' Remove any characters beyond the exe name:
            iPos = InStr(1, sPath, sEXE, CompareMethod.Text)
            sPath = Microsoft.VisualBasic.Left(sPath, iPos + Len(sEXE) - 1)

            Return Trim(sPath)
        Catch ex As Exception
            Return ""
            StatusBar.Text = "Error During Get Office Path"
        End Try
    End Function

    Private Sub Print_Report()
        'Prints the "Summary of Sales by Year" report in Northwind.mdb.

        ' Enable an error handler for this procedure:
        On Error GoTo ErrorHandler

        Dim oAccess As Access.Application
        Dim sDBPath As String 'path to Northwind.mdb
        Dim sReport As String 'name of report to print

        sReport = "Summary of Sales by Year"

        ' Shutdown any running instances of Access
        Dim MyProcesses() As Process
        Dim MyProcessID As Integer
        Dim MyProcessName As String
        MyProcesses = Process.GetProcessesByName("MSACCESS")
        Dim ProcessCnt As Integer = MyProcesses.Length
        Dim cntr As Long
        Do While cntr < ProcessCnt
            MyProcessID = MyProcesses(cntr).Id
            MyProcessName = MyProcesses(cntr).ProcessName
            If MyProcessName = "MSACCESS" Then
                MyProcesses(cntr).Kill()
            End If
            cntr += 1
        Loop

        ' Start a new instance of Access for automation:
        oAccess = New Access.ApplicationClass



        '  MyProcessID = MyProcesses(0).Id

        ' Determine the path to Northwind.mdb:
        sDBPath = oAccess.SysCmd(Action:=Access.AcSysCmdAction.acSysCmdAccessDir)
        sDBPath = sDBPath & "Samples\Northwind.mdb"

        ' Open Northwind.mdb in shared mode:
        oAccess.OpenCurrentDatabase(filepath:=sDBPath, Exclusive:=False)

        ' Select the report name in the database window and give focus
        ' to the database window:
        oAccess.DoCmd.SelectObject(ObjectType:=Access.AcObjectType.acReport,
            ObjectName:=sReport, InDatabaseWindow:=True)

        ' Output Report to Disk
        Dim outputfilename As String = "C:\Online Reports\Myreport.rtf"
        '        oAccess.DoCmd.OpenReport(ReportName:=sReport, _
        '                 View:=Access.AcView.acViewPreview)
        '      oAccess.DoCmd.Maximize()
        '   oAccess.CommandBars("Menu Bar").Enabled = False
        '  oAccess.CommandBars("Print Preview").Enabled = False
        ' oAccess.CommandBars("Print Preview Popup").Enabled = False
        If System.IO.File.Exists(outputfilename) Then
            System.IO.File.Delete(outputfilename)
            System.Threading.Thread.Sleep(1000)

        End If
        System.Threading.Thread.Sleep(1000)
        '    MsgBox(System.Windows.Forms.DataFormats.Rtf)
        '    oAccess.DoCmd.OutputTo(Access.AcOutputObjectType.acOutputReport, sReport, System.Windows.Forms.DataFormats.Rtf, outputfilename, False)

        ' Print the report:
        oAccess.DoCmd.OpenReport(ReportName:=sReport,
             View:=Access.AcView.acViewNormal)
        System.Threading.Thread.Sleep(1000)
        oAccess.DoCmd.Close(Access.AcObjectType.acReport, sReport, Access.AcCloseSave.acSaveNo)


        oAccess.Quit()

Cleanup:
        ' Quit Access and release object:
        On Error Resume Next
        oAccess.Quit(Option:=Access.AcQuitOption.acQuitSaveNone)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oAccess)
        oAccess = Nothing

        MyProcesses = Process.GetProcessesByName("AccessAutomation")
        MyProcessName = MyProcesses(0).ProcessName
        MyProcesses(0).Kill()



        Exit Sub
ErrorHandler:
        MsgBox(Err.Number & ": " & Err.Description,
            MsgBoxStyle.MsgBoxSetForeground, "Error Handler")
        ' Try to quit Access due to an unexpected error:
        Resume Cleanup
    End Sub

    Private Sub Preview_Report()
        'Previews the "Summary of Sales by Year" report in Northwind.mdb.

        ' Enable an error handler for this procedure:
        On Error GoTo ErrorHandler

        Dim oAccess As Access.Application
        Dim oForm As Object

        Dim sDBPath As String 'path to Northwind.mdb
        Dim sReport As String 'name of report to preview

        sReport = "Summary of Sales by Year"

        ' Start a new instance of Access for automation:
        oAccess = New Access.ApplicationClass

        ' Make sure Access is visible:
        If Not oAccess.Visible Then oAccess.Visible = True

        ' Determine the path to Northwind.mdb:
        sDBPath = oAccess.SysCmd(Action:=Access.AcSysCmdAction.acSysCmdAccessDir)
        sDBPath = sDBPath & "Samples\Northwind.mdb"

        ' Open Northwind.mdb in shared mode:
        oAccess.OpenCurrentDatabase(filepath:=sDBPath, Exclusive:=False)

        ' Close any forms that Northwind may have opened:
        For Each oForm In oAccess.Forms
            oAccess.DoCmd.Close(ObjectType:=Access.AcObjectType.acForm,
                ObjectName:=oForm.Name,
                Save:=Access.AcCloseSave.acSaveNo)
        Next
        If Not oForm Is Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
        End If
        oForm = Nothing

        ' Select the report name in the database window and give focus
        ' to the database window:
        oAccess.DoCmd.SelectObject(ObjectType:=Access.AcObjectType.acReport,
            ObjectName:=sReport, InDatabaseWindow:=True)

        ' Maximize the Access window:
        oAccess.RunCommand(Command:=Access.AcCommand.acCmdAppMaximize)

        ' Preview the report:
        oAccess.DoCmd.OpenReport(ReportName:=sReport,
            View:=Access.AcView.acViewPreview)

        ' Maximize the report window:
        oAccess.DoCmd.Maximize()

        ' Hide Access menu bar:
        oAccess.CommandBars("Menu Bar").Enabled = False

        ' Hide Report's Print Preview menu bar:
        oAccess.CommandBars("Print Preview").Enabled = False

        ' Hide Report's right-click popup menu:
        oAccess.CommandBars("Print Preview Popup").Enabled = False

        ' Save Report to Disk
        '  Shell("Kill c:\myreport.rtf")

        '    oAccess.DoCmd.OutputTo(Access.AcOutputObjectType.acOutputReport, sReport, System.Windows.Forms.DataFormats.Rtf, "c:\myreport.rtf", True)
        ' Close Report
        oAccess.DoCmd.Close(Access.AcObjectType.acReport, sReport, Access.AcCloseSave.acSaveNo)


        ' Release Application object and allow Access to be closed by user:
        If Not oAccess.UserControl Then oAccess.UserControl = True
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oAccess)
        oAccess = Nothing

        Exit Sub
ErrorCleanup:
        ' Try to quit Access due to an unexpected error:
        On Error Resume Next
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
        oForm = Nothing
        oAccess.Quit(Option:=Access.AcQuitOption.acQuitSaveNone)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oAccess)
        oAccess = Nothing

        Exit Sub
ErrorHandler:
        MsgBox(Err.Number & ": " & Err.Description,
            MsgBoxStyle.MsgBoxSetForeground, "Error Handler")
        Resume ErrorCleanup
    End Sub

    Private Sub Show_Form()
        'Shows the "Customer Labels Dialog" form in Northwind.mdb
        'and manipulates controls on the form.

        ' Enable an error handler for this procedure:
        On Error GoTo ErrorHandler

        Dim oAccess As Access.Application
        Dim oForm As Access.Form
        Dim oCtls As Access.Controls
        Dim oCtl As Access.Control
        Dim sDBPath As String 'path to Northwind.mdb
        Dim sForm As String 'name of form to show

        sForm = "Customer Labels Dialog"

        ' Start a new instance of Access for automation:
        oAccess = New Access.ApplicationClass

        ' Make sure Access is visible:
        If Not oAccess.Visible Then oAccess.Visible = True

        ' Determine the path to Northwind.mdb:
        sDBPath = oAccess.SysCmd(Action:=Access.AcSysCmdAction.acSysCmdAccessDir)
        sDBPath = sDBPath & "Samples\Northwind.mdb"

        ' Open Northwind.mdb in shared mode:
        oAccess.OpenCurrentDatabase(filepath:=sDBPath, Exclusive:=False)

        ' Close any forms that Northwind may have opened:
        For Each oForm In oAccess.Forms
            oAccess.DoCmd.Close(ObjectType:=Access.AcObjectType.acForm,
                ObjectName:=oForm.Name,
                Save:=Access.AcCloseSave.acSaveNo)
        Next
        If Not oForm Is Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
        End If
        oForm = Nothing

        ' Select the form name in the database window and give focus
        ' to the database window:
        oAccess.DoCmd.SelectObject(ObjectType:=Access.AcObjectType.acForm,
            ObjectName:=sForm, InDatabaseWindow:=True)

        ' Show the form:
        oAccess.DoCmd.OpenForm(FormName:=sForm,
            View:=Access.AcFormView.acNormal)

        ' Use Controls collection to edit the form:
        oForm = oAccess.Forms(sForm)
        oCtls = oForm.Controls

        ' Set PrintLabelsFor option group to Specific Country:
        oCtl = oCtls.Item("PrintLabelsFor")
        oCtl.Value = 2 'second option in option group
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oCtl)
        oCtl = Nothing

        ' Put USA in the SelectCountry combo box:
        oCtl = oCtls.Item("SelectCountry")
        oCtl.Enabled = True
        oCtl.SetFocus()
        oCtl.Value = "USA"
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oCtl)
        oCtl = Nothing

        ' Hide the Database Window:
        oAccess.DoCmd.SelectObject(ObjectType:=Access.AcObjectType.acForm,
            ObjectName:=sForm, InDatabaseWindow:=True)
        oAccess.RunCommand(Command:=Access.AcCommand.acCmdWindowHide)

        ' Set focus back to the form:
        oForm.SetFocus()

        ' Release Controls and Form objects:
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oCtls)
        oCtls = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
        oForm = Nothing

        ' Release Application object and allow Access to be closed by user:
        If Not oAccess.UserControl Then oAccess.UserControl = True
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oAccess)
        oAccess = Nothing

        Exit Sub
ErrorCleanup:
        ' Try to quit Access due to an unexpected error:
        On Error Resume Next
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oCtl)
        oCtl = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oCtls)
        oCtls = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
        oForm = Nothing
        oAccess.Quit(Option:=Access.AcQuitOption.acQuitSaveNone)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oAccess)
        oAccess = Nothing

        Exit Sub
ErrorHandler:
        MsgBox(Err.Number & ": " & Err.Description,
            MsgBoxStyle.MsgBoxSetForeground, "Error Handler")
        Resume ErrorCleanup
    End Sub
    Private Function CompactRepair_SecureDB(ByVal sdbpath As String, ByVal mdwpath As String, ByVal suser As String, ByVal Spwd As String) As Boolean
        Dim sAccPath As String
        Dim scmdline As String
        Dim iSleepTime As Long
        Dim enuWindowStyle As Microsoft.VisualBasic.AppWinStyle
        enuWindowStyle = AppWinStyle.MinimizedFocus
        Try

            ' Make sure specified database (sDBPath) exists:
            If Not System.IO.File.Exists(sdbpath) Then
                MsgBox("Can't find the file '" & sdbpath & "'",
                    MsgBoxStyle.MsgBoxSetForeground)
                Return False
            End If
            sAccPath = "C:\Program Files (x86)\Microsoft Office\Office14\MSACCESS.EXE"
            scmdline = "/wrkgrp " & mdwpath
            Shell(PathName:=sAccPath & " " & scmdline & " " & sdbpath & " /repair /user " & suser & " /pwd " & Spwd, Style:=enuWindowStyle)

            ' Pause to allow database to open:
            '    iSleepTime = 6000
            iSleepTime = 2000
            System.Threading.Thread.Sleep(iSleepTime)
            '    AppActivate(Title:=myprocessname)
            ' Obtain Application object of the instance of Access
            ' that has the database open:
            '  MessageBox.Show("Before GetObject Command")

            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function
    Dim oAccess As Access.Application
    Private Function OpenAccessApplication(ByVal sProgPath As String, ByVal sdbpath As String, startupparms As String, mdwpath As String, ByVal suser As String, ByVal spwd As String) As Access.Application
        Try
            Dim oAccess As Access.Application
            Dim sCommand As String
            If mdwpath <> "" Then
                sCommand = sProgPath + " " + sdbpath + " /wrkgrp " + mdwpath + " /user " + suser + " /pwd " + spwd
            Else
             '   sCommand = sProgPath + " " + sdbpath + " /user AUTOREPORTS"
                sCommand = sProgPath + " " + sdbpath
            End If
            Dim dummy As Long
            dummy = Shell(sCommand, AppWinStyle.MinimizedFocus, Wait:=True, Timeout:=5000)

            oAccess = GetObject(sdbpath)


            '   oAccess = ShellGetDB(sdbpath, " /wrkgrp " & mdwpath & " /user " & suser & " /pwd " & spwd, AppWinStyle.MinimizedFocus, 3000)
            Return oAccess
        Catch ex As Exception
            Return Nothing
        End Try

    End Function
    Private Function CloseAccessApplication(oAccess As Access.Application) As Boolean
        Try
            SetUIText_ThreadSafe(Me.StatusBar, "Report Saved to Local Disk")
            System.Threading.Thread.Sleep(1000)
            '   oAccess.DoCmd.Close(Access.AcObjectType.acReport, sreport, Access.AcCloseSave.acSaveNo)
            '   oAccess.DoCmd.Close(Access.AcObjectType.acQuery, sreport, Access.AcCloseSave.acSaveNo)




            SetUIText_ThreadSafe(Me.StatusBar, "Report/Query Object Closed")
            '  
            '  This next line was commented out because it seems to be a property you can't set.
            '    If Not oAccess.UserControl Then oAccess.UserControl = True

            Try
                oAccess.DoCmd.Quit()

                While (System.Runtime.InteropServices.Marshal.ReleaseComObject(oAccess) > 0)
                End While


            Catch
                SetUIText_ThreadSafe(Me.StatusBar, "Error during Shutdown")
            Finally
                oAccess = Nothing
            End Try

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Private Function GenerateOutput(ByVal JobID As Integer, ByVal sreport As String, ByVal scriteria As String, ByVal outputdirectory As String, ByVal outputfilename As String, ByVal outputformat As String) As Boolean
        'Shows how to automate Access when user-level
        'security is enabled and you wish to avoid the Logon
        'dialog asking for user name and password. In this 
        'example we're assuming default security so we simply
        'pass the Admin user with a blank password to print the 
        '"Summary of Sales by Year" report in Northwind.mdb.
        Dim cdi As New CDIntfEx.CDIntfEx
        Dim localfilepath As String
        Try
            If oAccess Is Nothing Then
                Return False
            End If

            localfilepath = outputdirectory & outputfilename

            If outputformat = "XLS" Then
                localfilepath = Replace(localfilepath, ".OUT", ".XLS")
                oAccess.DoCmd.SelectObject(ObjectType:=Access.AcObjectType.acQuery, ObjectName:=sreport, InDatabaseWindow:=True)
                oAccess.DoCmd.OpenQuery(sreport, View:=Access.AcView.acViewNormal)
                oAccess.DoCmd.Maximize()
                SetUIText_ThreadSafe(Me.StatusBar, "Running Query")
            Else
                If outputformat = "CSV" Then
                    localfilepath = Replace(localfilepath, ".OUT", ".CSV")
                    '   oAccess.DoCmd.SelectObject(ObjectType:=Access.AcObjectType.acQuery, ObjectName:=sreport, InDatabaseWindow:=True)
                    '   oAccess.DoCmd.OpenQuery(sreport, View:=Access.AcView.acViewNormal)
                    '   oAccess.DoCmd.Maximize()
                    '   SetUIText_ThreadSafe(Me.StatusBar, "Opening Query for Export")

                Else
                     If System.IO.File.Exists(localfilepath) Then
                        System.IO.File.Delete(localfilepath)
                        System.Threading.Thread.Sleep(1000)
                        SetUIText_ThreadSafe(Me.StatusBar, "Report Existed - Deleted")
                     End If



                    oAccess.DoCmd.SelectObject(ObjectType:=Access.AcObjectType.acReport, ObjectName:=sreport, InDatabaseWindow:=True)
                    oAccess.DoCmd.OpenReport(ReportName:=sreport, View:=Access.AcView.acViewPreview, WhereCondition:=scriteria)


                   ' Dim AcFileFormatPDF As Object = Nothing
                   ' oAccess.DoCmd.OutputTo(AcOutputObjectType.acOutputReport, sreport, AcFileFormatPDF, localfilepath,False,,,AcExportQuality.acExportQualityPrint)
                    ' 2501 ERR IS NO DATA
                    oAccess.DoCmd.Maximize()
                    oAccess.Reports(sreport).FilterOn = True
                    SetUIText_ThreadSafe(Me.StatusBar, "Running Report ...(" + sreport + ") " + scriteria)
                End If
            End If
            If System.IO.File.Exists(localfilepath) Then
                System.IO.File.Delete(localfilepath)
                System.Threading.Thread.Sleep(1000)
                SetUIText_ThreadSafe(Me.StatusBar, "Report Existed - Deleted")
            End If

            System.Threading.Thread.Sleep(1000)

            Select Case outputformat
                Case Is = "PDF"
                    If System.IO.File.Exists(localfilepath) then
                        System.IO.File.Delete(localfilepath)
                        System.Threading.Thread.Sleep(1000)
                           SetUIText_ThreadSafe(Me.StatusBar, "Report Existed - Deleted")
                    End If    
                    Const acFormatPDF As String = "PDF Format(*.pdf)"
                    oAccess.DoCmd.OutputTo(Access.AcOutputObjectType.acOutputReport,sreport,acFormatPDF,localfilepath,,,,AcExportQuality.acExportQualityPrint)
                Case Is = "PDFAmyuni"
                    ' Save as a PDF
                    Try
                        '                        oAccess.DoCmd.OutputTo(Access.AcOutputObjectType.acOutputReport, sreport, System.Windows.Forms.DataFormats.Rtf, localfilepath, False)

                        cdi.DriverInit("Amyuni PDF Converter")
                        cdi.DefaultDirectory = outputdirectory
                        cdi.DefaultFileName = localfilepath
                        cdi.FileNameOptionsEx = 3
                        cdi.SetDefaultPrinter()
                        cdi.EnablePrinter("Sensible Solutions Inc.", "07EFCDAB01000100FDF4C119DEED385CAED59A31B52748A879E4800C076E69A111BC87B4941C9C596C1D7E65D6C48370FF2A0CF1C1A8")
                        oAccess.DoCmd.PrintOut()
                        cdi.RestoreDefaultPrinter()
                    Catch ex As Exception
                        Return False
                    End Try
                Case Is = "PDFx"
                    Try
                        cdi.DriverInit("CutePDF Writer")
                        cdi.DefaultDirectory = outputdirectory
                        cdi.DefaultFileName = localfilepath
                        cdi.FileNameOptionsEx = 3
                        cdi.SetDefaultPrinter()
                        cdi.EnablePrinter("Sensible Solutions Inc.", "")
                        ' cdi.EnablePrinter("Sensible Solutions Inc.", "07EFCDAB01000100FDF4C119DEED385CAED59A31B52748A879E4800C076E69A111BC87B4941C9C596C1D7E65D6C48370FF2A0CF1C1A8")
                        oAccess.DoCmd.PrintOut()
                        cdi.RestoreDefaultPrinter()



                    Catch ex As Exception

                    End Try
                Case Is = "rtf"
                    '  oAccess.DoCmd.OutputTo(Access.AcOutputObjectType.acOutputReport, sreport, System.Windows.Forms.DataFormats.Rtf, outputfilepath, False)
                Case Is = "snp"
                    '     oAccess.DoCmd.OutputTo(Access.AcOutputObjectType.acOutputReport, sreport, System.Windows.Forms.DataFormats.Rtf, outputfilepath, False)
                Case Is = "html"
                    '     oAccess.DoCmd.OutputTo(Access.AcOutputObjectType.acOutputReport, sreport, System.Windows.Forms.DataFormats.Html, outputfilepath, False)
                Case Is = "XLS"
                    '   oAccess.DoCmd.OutputTo(Access.AcOutputObjectType.acOutputQuery, sreport, System.Windows.Forms.DataFormats.Text, outputfilepath, False)
                    '  oAccess.DoCmd.OutputTo(Access.AcOutputObjectType.acOutputQuery, sreport, System.Windows.Forms.DataFormats.CommaSeparatedValue, outputfilename, False)
                    '  outputfilepath = "c:   \windows\system32\AR000289.xls"

                    oAccess.DoCmd.SetWarnings(vbFalse)
                    oAccess.DoCmd.TransferSpreadsheet(Access.AcDataTransferType.acExport, Access.AcSpreadSheetType.acSpreadsheetTypeExcel8, sreport, localfilepath, True, , True)

                Case Is = "CSV"
                    oAccess.DoCmd.SetWarnings(False)
                    oAccess.DoCmd.RunSavedImportExport(sreport)
                    oAccess.DoCmd.SetWarnings(True)

                    '       oAccess.DoCmd.TransferText(Access.AcDataTransferType.acExport, "PartList Export Specification", "PcWtLookupTable", "PARTLIST.CSV", True)
                    '       oAccess.DoCmd.OutputTo(Access.AcOutputObjectType.acOutputQuery, sreport, System.Windows.Forms.DataFormats.CommaSeparatedValue, outputfilename, False, "Export-PartList", , Access.AcExportQuality.acExportQualityScreen)
                    '  DoCmd.TransferText acExportDelim, "PartList Export Specification", "PcWtLookupTable", "C:\PARTLIST.CSV", True
            End Select

            SetUIText_ThreadSafe(Me.StatusBar, "Report Saved to Local Disk to " + localfilepath)
            System.Threading.Thread.Sleep(1000)
            oAccess.DoCmd.Close(Access.AcObjectType.acReport, sreport, Access.AcCloseSave.acSaveNo)
            oAccess.DoCmd.Close(Access.AcObjectType.acQuery, sreport, Access.AcCloseSave.acSaveNo)
            Return True
        Catch ex As Exception
            SetUIText_ThreadSafe(Me.StatusBar, ex.Message & " Error Occured During Generate Output")
            Return False
        End Try
    End Function


    Private Sub Print_Report_Security(ByVal sdbpath As String, ByVal sreport As String, ByVal scriteria As String, ByVal suser As String, ByVal spwd As String, ByVal outputdirectory As String, ByVal outputfilename As String, ByVal stroutputformat As String)
        'Shows how to automate Access when user-level
        'security is enabled and you wish to avoid the Logon
        'dialog asking for user name and password. In this 
        'example we're assuming default security so we simply
        'pass the Admin user with a blank password to print the 
        '"Summary of Sales by Year" report in Northwind.mdb.

        'Dim cdi As New CDIntfEx.CDIntfEx


        Dim oAccess As Access.Application
        Try
            If Not System.IO.File.Exists(sdbpath) Then
                MsgBox("Can't find the file '" & sdbpath & "'",
                    MsgBoxStyle.MsgBoxSetForeground)
                Exit Sub
            End If
            oAccess = ShellGetDB(sdbpath, " /nostartup /wrkgrp t:\hts\hts.mdw /user " & suser & " /pwd " & spwd, AppWinStyle.MinimizedFocus, 3000)

            If stroutputformat = "XLS" Then
                oAccess.DoCmd.SelectObject(ObjectType:=Access.AcObjectType.acQuery, ObjectName:=sreport, InDatabaseWindow:=True)
                oAccess.DoCmd.OpenQuery(sreport, View:=Access.AcView.acViewNormal)
                oAccess.DoCmd.Maximize()
                SetUIText_ThreadSafe(Me.StatusBar, "Query Window Maximized")

            Else
                oAccess.DoCmd.SelectObject(ObjectType:=Access.AcObjectType.acReport, ObjectName:=sreport, InDatabaseWindow:=True)

                oAccess.DoCmd.OpenReport(ReportName:=sreport, View:=Access.AcView.acViewPreview, WhereCondition:=scriteria)
                oAccess.DoCmd.Maximize()
                oAccess.Reports(sreport).FilterOn = True
                SetUIText_ThreadSafe(Me.StatusBar, "Report Window Maximized")

            End If

            '   oAccess.CommandBars("Menu Bar").Enabled = False
            '  oAccess.CommandBars("Print Preview").Enabled = False
            ' oAccess.CommandBars("Print Preview Popup").Enabled = False
            Dim outputfilepath As String = outputdirectory & outputfilename
            If System.IO.File.Exists(outputfilepath) Then
                System.IO.File.Delete(outputfilepath)
                System.Threading.Thread.Sleep(1000)
                SetUIText_ThreadSafe(Me.StatusBar, "Report Existed - Deleted")
            End If

            System.Threading.Thread.Sleep(1000)

            Select Case stroutputformat
                Case Is = "PDF"
                    ' Save as a PDF
                    Try
                        'cdi.DriverInit("Amyuni PDF Converter")
                        'cdi.DefaultDirectory = outputdirectory
                        'cdi.DefaultFileName = outputdirectory & outputfilename

                        ''     cdi.DefaultDirectory = "c:\"
                        ''   cdi.DefaultFileName = "c:\hts\reports\TEST.PDF"
                        'cdi.FileNameOptionsEx = 3
                        'cdi.SetDefaultPrinter()
                        'cdi.EnablePrinter("Sensible Solutions Inc.", "07EFCDAB01000100FDF4C119DEED385CAED59A31B52748A879E4800C076E69A111BC87B4941C9C596C1D7E65D6C48370FF2A0CF1C1A8")
                        ''   cdi.EnablePrinter("Amyuni PDF Converter Evaluation", "07EFCDAB0100010025AFF18045B8441306C5739F7DC52654D393BA9CECBA2ADE79E3762A65FFC354528A5F4A5811BE3204A0A439F5BA")
                        'oAccess.DoCmd.PrintOut()
                        'cdi.RestoreDefaultPrinter()
                        ''       cdi.FileNameOptions = 0

                    Catch ex As Exception

                    End Try
                Case Is = "rtf"
                    '    oAccess.DoCmd.OutputTo(Access.AcOutputObjectType.acOutputReport, sreport, System.Windows.Forms.DataFormats.Rtf, outputfilepath, False)
                Case Is = "snp"
                    '     oAccess.DoCmd.OutputTo(Access.AcOutputObjectType.acOutputReport, sreport, System.Windows.Forms.DataFormats.Rtf, outputfilepath, False)
                Case Is = "html"
                    '     oAccess.DoCmd.OutputTo(Access.AcOutputObjectType.acOutputReport, sreport, System.Windows.Forms.DataFormats.Html, outputfilepath, False)
                Case Is = "XLS"
                    '   oAccess.DoCmd.OutputTo(Access.AcOutputObjectType.acOutputQuery, sreport, System.Windows.Forms.DataFormats.Text, outputfilepath, False)
                    '  oAccess.DoCmd.OutputTo(Access.AcOutputObjectType.acOutputQuery, sreport, System.Windows.Forms.DataFormats.CommaSeparatedValue, outputfilename, False)
                    '  outputfilepath = "c:\windows\system32\AR000289.xls"
                    oAccess.DoCmd.TransferSpreadsheet(Access.AcDataTransferType.acExport, Access.AcSpreadSheetType.acSpreadsheetTypeExcel8, sreport, outputfilepath, True, , True)
            End Select

            SetUIText_ThreadSafe(Me.StatusBar, "Report Saved to Local Disk")
            System.Threading.Thread.Sleep(1000)
            oAccess.DoCmd.Close(Access.AcObjectType.acReport, sreport, Access.AcCloseSave.acSaveNo)
            oAccess.DoCmd.Close(Access.AcObjectType.acQuery, sreport, Access.AcCloseSave.acSaveNo)
            StatusBar.Text = "Report/Query Object Closed"
            If Not oAccess.UserControl Then oAccess.UserControl = True
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oAccess)
            oAccess = Nothing
            System.Threading.Thread.Sleep(1000)
        Catch ex As Exception
            SetUIText_ThreadSafe(Me.StatusBar, ex.Message & " Detected in Print Report Security Catch Block")

        Finally
            ' Quit Access and release object:
            System.Threading.Thread.Sleep(2000)
            StatusBar.Text = "before quit in Finally block"
            oAccess.Quit(Option:=Access.AcQuitOption.acQuitSaveNone)
            StatusBar.Text = "after quit in Finally block"
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oAccess)
            oAccess = Nothing

            System.Threading.Thread.Sleep(2000)
        End Try
    End Sub

    Private Sub Preview_Report_Runtime()
        'Shows how to automate the Access Runtime to preview
        'the "Summary of Sales by Year" report in Northwind.mdb.

        ' Enable an error handler for this procedure:
        On Error GoTo ErrorHandler

        Dim oAccess As Access.Application
        Dim oForm As Access.Form
        Dim sDBPath As String 'path to Northwind.mdb
        Dim sReport As String 'name of report to preview

        sReport = "Summary of Sales by Year"

        ' Determine the path to Northwind.mdb:
        sDBPath = GetOfficeAppPath("Access.Application", "msaccess.exe")
        If sDBPath = "" Then
            MsgBox("Can't determine path to msaccess.exe",
                MsgBoxStyle.MsgBoxSetForeground)
            Exit Sub
        End If
        sDBPath = Microsoft.VisualBasic.Left(sDBPath,
            Len(sDBPath) - Len("msaccess.exe")) & "Samples\Northwind.mdb"
        If Not System.IO.File.Exists(sDBPath) Then
            MsgBox("Can't find the file '" & sDBPath & "'",
                MsgBoxStyle.MsgBoxSetForeground)
            Exit Sub
        End If

        ' Start a new instance of Access. If the retail
        ' version of Access is not installed, and only the
        ' Access Runtime is installed, launches a new instance
        ' of the Access Runtime (/runtime switch is optional):
        oAccess = ShellGetDB(sDBPath, "/runtime")
        'or
        'oAccess = ShellGetApp(Chr(34) & sDBPath & Chr(34) & " /runtime")

        ' Make sure Access is visible:
        If Not oAccess.Visible Then oAccess.Visible = True

        ' Close any forms that Northwind may have opened:
        For Each oForm In oAccess.Forms
            oAccess.DoCmd.Close(ObjectType:=Access.AcObjectType.acForm,
                ObjectName:=oForm.Name,
                Save:=Access.AcCloseSave.acSaveNo)
        Next
        If Not oForm Is Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
        End If
        oForm = Nothing

        ' Select the report name in the database window and give focus
        ' to the database window:
        oAccess.DoCmd.SelectObject(ObjectType:=Access.AcObjectType.acReport,
            ObjectName:=sReport, InDatabaseWindow:=True)

        ' Maximize the Access window:
        oAccess.RunCommand(Command:=Access.AcCommand.acCmdAppMaximize)

        ' Preview the report:
        oAccess.DoCmd.OpenReport(ReportName:=sReport,
            View:=Access.AcView.acViewPreview)

        ' Maximize the report window:
        oAccess.DoCmd.Maximize()

        ' Hide Access menu bar:
        oAccess.CommandBars("Menu Bar").Enabled = False

        ' Release Application object and allow Access to be closed by user:
        If Not oAccess.UserControl Then oAccess.UserControl = True
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oAccess)
        oAccess = Nothing

        Exit Sub
ErrorCleanup:
        ' Try to quit Access due to an unexpected error:
        On Error Resume Next
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
        oForm = Nothing
        oAccess.Quit(Option:=Access.AcQuitOption.acQuitSaveNone)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oAccess)
        oAccess = Nothing

        Exit Sub
ErrorHandler:
        MsgBox(Err.Number & ": " & Err.Description,
            MsgBoxStyle.MsgBoxSetForeground, "Error Handler")
        Resume ErrorCleanup
    End Sub

    Private Sub Schedule_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        LoadSettings()

        '  Setup Timer / Background Worker
        Timer1.Interval = 1000
        Timer1.Enabled = False

        StatusBar.Text = "Retrieving Current Data from Web Service"
        refreshdata()


        If My.Settings.AutoStartOnLoad = True Then
            Online.Checked = True
            StopButton.Enabled = True
            StartButton.Enabled = False
        Else
            Online.Checked = False
            StopButton.Enabled = False
            StartButton.Enabled = True
        End If

        Timer1.Enabled = True

        Running.Checked = False
        Distributing.Checked = False
        LastCheck.Text = DateTime.Now.ToString
        LastDistributeCheck.Text = DateTime.Now.ToString
        LastMessageRun.Text = DateTime.Now.ToString

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub ToolBar1_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs)
        Select Case e.Button.Text
            Case "Start"
                ' Enable Timer

                '    Me.StartButton.Enabled = True
                Me.Online.Checked = True
                '  Reports will start running at next Scheduled Time based on App Setting
            Case "Stop"
                Me.Online.Checked = False
                StopReportGenerator()
                StopReportDistributor()
                '  Report Generator will shut down after next report is completed
            Case "Run Now"
                StartReportGenerator()
                StartReportDistributor()
            Case "Settings"
                Dim Settings As New Settings()
                Settings.ShowDialog()
                LoadSettings()
            Case "Refresh"
                Try
                    refreshdata()
                    StatusBar.Text = "Data Refreshed at " & Now.ToString
                Catch EX As Exception
                    StatusBar.Text = "An Error Occured Retrieving Data from the Web Service" + EX.ToString
                End Try
            Case "TestWS"
                ' Todo Add Method to WCF Service to test connection
                '   Me.StatusBar.Text = ws.HelloWorld

                '     Me.StatusBar.Text = MsgBox(ar.WhoAmI)
            Case "Settings"
        End Select

    End Sub

    Private Sub StartReportGenerator()
        '    Me.StopStart.Enabled = False
        '    Me.StopButton.Enabled = True
        Me.Running.Checked = True
        BGReportGenerator.RunWorkerAsync()
        StatusBar.Text = "Report Generator Started at " + Now.ToShortTimeString
    End Sub

    Private Sub StartReportDistributor()
        Me.Distributing.Checked = True
        BGReportDistributor.RunWorkerAsync()
        StatusBar.Text = "Report Distributor Started at " + Now.ToShortTimeString
    End Sub
    Private Sub StopReportDistributor()
        If BGReportDistributor.IsBusy Then
            If BGReportDistributor.WorkerSupportsCancellation Then
                BGReportDistributor.CancelAsync()
            Else
                Distributing.Checked = False
            End If
        Else
            Distributing.Checked = False
        End If

    End Sub
    Private Sub StopReportGenerator()
        If BGReportGenerator.IsBusy Then
            If BGReportGenerator.WorkerSupportsCancellation Then
                BGReportGenerator.CancelAsync()
            End If
            '        Me.StartButton.Enabled = True
            '        Me.StopButton.Enabled = False
        Else
            Running.Checked = False
        End If
    End Sub

    Private Sub StartMessenger()
        '    Me.StopStart.Enabled = False
        '    Me.StopButton.Enabled = True
        Me.SendingMessages.Checked = True
        BGMessenger.RunWorkerAsync()
        StatusBar.Text = "Messenger Generator Started at " + Now.ToShortTimeString
    End Sub

    Private Sub StopMessenger()
        If BGMessenger.IsBusy Then
            If BGMessenger.WorkerSupportsCancellation Then
                BGMessenger.CancelAsync()
            End If
            '        Me.StartButton.Enabled = True
            '        Me.StopButton.Enabled = False
        Else
            Running.Checked = False
        End If

    End Sub
    Private Sub DataGrid1_Navigate1(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs)
        Dim strCurrentJob As String
        '       strCurrentJob = DataView1(dgJobSchedule.CurrentRowIndex)("JOB").ToString
        '       dvHistory.RowFilter = "JOB = " & strCurrentJob
    End Sub


    Private Sub TreeView1_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterSelect
        Dim selecteditem As String
        Dim selectedview As String
        Try
            Select Case TreeView1.SelectedNode.Parent.Text
                Case "All Applications"
                    selecteditem = TreeView1.SelectedNode.Text
                    selectedview = "Application"
                Case "All Users"


            End Select
            Select Case selectedview
                Case Is = "Application"
                    Dim app As New ViewApplication
                    app.Show()
            End Select
        Catch

        End Try
    End Sub


    Private Sub Schedule_Closing1(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        '   Debug.WriteLine("Form Closing")
        ' Create a copy of the xml file for offline use if necessary
        ardata.WriteXml("ardata.xml")

        ' Persist session variables

        ' Use Reflection to find the location of the config file
        Dim asm As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly
        Dim strconfigloc As String
        strconfigloc = asm.Location

        ' The config file is located in the application's bin directory, remove the filename

        Dim strtemp As String
        strtemp = strconfigloc
        strtemp = System.IO.Path.GetDirectoryName(strconfigloc)

        ' Declare a FileInfo object for the config file
        Dim fileinfo As System.IO.FileInfo = New System.IO.FileInfo(strtemp & "\AccessAutomation.exe.config")

        ' Load the config file into the XML DOM
        Dim xmldocument As New System.Xml.XmlDocument
        xmldocument.Load(fileinfo.FullName)

        ' For Each Node / Reset Value to current setting which will be reset at startup
        Dim node As System.Xml.XmlNode
        For Each node In xmldocument.Item("configuration").Item("appSettings")
            ' Skip any comments
            If node.Name = "add" Then
                If node.Attributes.GetNamedItem("key").Value = "WorkOffline" Then
                    node.Attributes.GetNamedItem("value").Value = CType(Me.WorkOffline.Checked, String)
                End If
                If node.Attributes.GetNamedItem("key").Value = "StartOffline" Then
                    If Online.Checked Then
                        node.Attributes.GetNamedItem("value").Value = "False"
                    Else
                        node.Attributes.GetNamedItem("value").Value = "True"
                    End If
                End If
                If node.Attributes.GetNamedItem("key").Value = "Running" Then
                    node.Attributes.GetNamedItem("value").Value = CType(Me.Running.Checked, String)
                End If

            End If
        Next node

        ' Save the modified config file
        xmldocument.Save(fileinfo.FullName)

    End Sub

    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)

    End Sub

    Private Sub daHistory_RowUpdated(ByVal sender As System.Object, ByVal e As System.Data.OleDb.OleDbRowUpdatedEventArgs)

    End Sub

    Private Sub EmailTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EmailTest.Click
        Dim dl() As UserList
        dl = ws.GetJobDistributionList(311)
        SendEmailX(dl, "edolikian@htsmi.com", "Test Message", "Test Body Text", "")
        '   SendEmailMessagey("edolikian@ssitroy.com", "edolikian@htsmi.com", "Test Message", "Test Body Text", "")
    End Sub
    Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMS As Long)
    Private Sub Button2_Click(sender As Object, e As EventArgs)
        Dim accObj As Access.Application
        Dim application As String
        Dim dbs As String
        Dim workgroup As String
        Dim msg As String
        Dim user As String
        Dim password As String
        Dim ctries As Integer
        Dim x As Object
        '          "C:\Program Files (x86)\Microsoft Office\Office14\MSACCESS.EXE" /wrkgrp  t:\hts\hts.mdw  c:\hts\app2k3\htsapp2k3.mdb

        application = "C:\Program Files (x86)\Microsoft Office\Office14\MSACCESS.EXE"
        dbs = "c:\hts\app2k3\htsapp2k3.mdb"
        user = "autoreports"
        password = "4438"
        workgroup = "t:\hts\hts.mdw"
        x = Shell(application & " " & Chr(34) & dbs & Chr(34) & " /nostartup /user " & user & " /pwd " & password & " /wrkgrp " & Chr(34) & workgroup & Chr(34), vbMinimizedFocus)
        On Error GoTo WaitforAccess
        accObj = GetObject(, "Access.Application")
        '
        On Error GoTo 0

        msg = "Access is now open."
        MsgBox(msg, , "Success!")

        Dim pdfFileNameToStore As String
        pdfFileNameToStore = "c:\test.pdf"
        Dim acformatpdf As String = "PDF Format (*.pdf)"

        accObj.DoCmd.OutputTo(Access.AcOutputObjectType.acOutputReport, "Open Lot Printout", acformatpdf, pdfFileNameToStore, False, , , Access.AcExportQuality.acExportQualityPrint)


        ' Open a Report and Save as a PDF
        '     accObj.DoCmd.OpenReport("Open Lot Printout", Access.AcView.acViewPreview, "[CUST ID] = '0164'", Access.AcWindowMode.acWindowNormal)
        '      accObj.DoCmd.OutputTo(Access.AcOutputObjectType.acOutputReport, "Open Lot Printout", "PDF Format (*.pdf)", "C:\test.pdf", False, , , Access.AcExportQuality.acExportQualityPrint)

        accObj.DoCmd.OutputTo(Access.AcOutputObjectType.acOutputReport, "Open Lot Printout", acformatpdf, "c:\myreport.pdf", False, , , Access.AcExportQuality.acExportQualityPrint)

        '    accObj.DoCmd.OutputTo(Access.AcOutputObjectType.acOutputReport, "Open Lot Printout", "PDF Format (*.pdf)", "C:\Temp\ComputerCounts.pdf")



        accObj.CloseCurrentDatabase()
        accObj.Quit()

        accObj = Nothing
        MsgBox("All Done", vbMsgBoxSetForeground)

        Exit Sub

WaitforAccess:
        '   SetFocus()
        If ctries < 5 Then

            ctries = ctries + 1
            Sleep(500)
            Resume
        Else
            MsgBox("Access is taking too long. Process Ended", vbMsgBoxSetForeground)
        End If

    End Sub

    Private Sub RunReportsButton_Click(sender As Object, e As EventArgs) Handles RunReportsButton.Click
        RunOverdueReports("Ascending")
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs)

    End Sub


    Private Sub StatusBar_TextChanged(sender As Object, e As EventArgs) Handles StatusBar.TextChanged

    End Sub

    Friend myProcessArray As New ArrayList
    Private myProcess As Process


    Private Function getFileProcess(ByVal strFile As String) As ArrayList
        myProcessArray.Clear()

        Dim processes As Process() = Process.GetProcesses
        Dim i As Integer
        Try
            For i = 0 To processes.GetUpperBound(0) - 1
                myProcess = processes(i)
                If Not myProcess.HasExited And Not myProcess.ProcessName = "System" Then
                    Try
                        Dim modules As ProcessModuleCollection = myProcess.Modules
                        Dim j As Integer
                        For j = 0 To modules.Count - 1
                            If (modules.Item(j).FileName.ToLower.CompareTo(strFile.ToLower) = 0) Then
                                myProcessArray.Add(myProcess)
                                Exit For
                            End If
                        Next j
                    Catch ex As Exception
                        ' Msgbox(("Error: & ex.Message()
                    End Try
                End If
            Next i



        Catch ex As Exception
            Debug.WriteLine(ex.Message.ToString)

        End Try

        Return myProcessArray
    End Function





    Private Sub BGReportGenerator_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BGReportGenerator.DoWork


        '  Setup / Start Timer
        Dim CheckStart As Date
        Dim ws As New AutoReportsWCFService.ServiceClient
        Dim al() As AppsList
        Dim jl() As JobList
        Dim reportoutcome As Boolean
        Dim uploadoutcome As Boolean
        Dim jobcompletion As DateTime
        Dim archcontainer As String

        Dim strtempdirectory As String
        Dim strtempfilename As String
        Dim strtempextension As String
        Dim strtempfilepath As String
        Dim archfilename As String
        Dim strOutputFormat As String
        Dim ArchiveID As Integer
        Dim ctr As Integer
        '  Start Collection Cycle
        Try


            CheckStart = Now
            '     Timer1.Enabled = False
            al = ws.GetAppsList()

            '  Shut Down / Kill any instances of Access that are already running

            Dim pProcess() As Process = Process.GetProcessesByName("MSACCESS")
            For Each p As Process In pProcess
                If p.ProcessName = "MSACCESS" Then
                    '   For Each m As ProcessModule In p.Modules
                    '    Dim a As ArrayList = getFileProcess("C:\HTS\APP2K3\HTSAPP2K3.MDB")
                    '   For Each PP As Process In a
                    ' Debug.Print(PP.ProcessName & PP.SessionId)
                    ' Next
                    p.Kill()
                End If
            Next


            '  Process Each App / Look for Overdue Jobs / Run / Distribute / Mark as Sent
            For Each app As AppsList In al
                SetUIText_ThreadSafe(Me.StatusBar, "Checking App " + app.DESCRIPTION + " for Overdue Reports ....")
                jl = ws.GetOverdueJobsList(app.APPID, DateTime.Now.ToString())


                SetUIText_ThreadSafe(Me.StatusBar, "There are " + jl.Length.ToString + " Overdue Jobs to Run in " + app.DESCRIPTION + " at " + DateTime.Now().ToString())

                If jl.Length > 0 Then
                    '  Compact / Repair App to make sure it is not corrupted 
                    'If CompactRepair_SecureDB(app.STARTUP_DIRECTORY + app.STARTUP_FILENAME, app.MDW_DIRECTORY + app.MDW_FILENAME, app.USERNAME, app.PASSWORD) = True Then
                    '    Thread.Sleep(2000)
                    'Else
                    '    GoTo NextApp
                    'End If
                    '  Open App
                    Dim sAccPath As String
                    '   sAccPath = "C:\Program Files (x86)\Microsoft Office\Office14\MSACCESS.EXE"
                    sAccPath = "C:\Program Files (x86)\Microsoft Office\root\Office16\MSACCESS.EXE"

                    If app.APPID = 3 Then
                        sAccPath = "C:\Program Files (x86)\Microsoft Office\root\Office16\MSACCESS.EXE"
                        sAccPath = "C:\Program Files (x86)\Microsoft Office\Office16\MSACCESS.EXE"
                        app.STARTUP_DIRECTORY = "c:\HTS\APP2K3\"
                        app.STARTUP_FILENAME = "HTSAPP2K3.MDB"
                        app.MDW_DIRECTORY = ""
                        app.MDW_FILENAME = ""
                        '   app.USERNAME = "AUTOREPORTS"
                        '   app.PASSWORD = "4438"
                     '   app.USERNAME = "AUTOREPORTS"
                     '   app.PASSWORD = ""
                    End If

                    Try
                        '     If app.APPID <> 3 Then
                        oAccess = OpenAccessApplication(sAccPath, app.STARTUP_DIRECTORY + app.STARTUP_FILENAME, app.STARTUP_OPTIONS, app.MDW_DIRECTORY + app.MDW_FILENAME, app.USERNAME, app.PASSWORD)
                        '     oAccess = OpenAccessApplication(sAccPath,)
                        '   Else
                        '   oAccess = OpenAccessApplication(sAccPath, "t:\HTS\APP2K3\HTSMST2013.ACCDB", "", "", "", "")
                        '   End If
                    Catch
                        oAccess = Nothing
                    End Try

                    If Not oAccess Is Nothing Then
                        ctr = 0
                        For Each OverdueJob As JobList In jl

                            ' Check if Cancel Requested / Stop Running Reports
                            If BGReportGenerator.CancellationPending Or ctr >= 15 Then
                                e.Cancel = True
                                Exit For
                            End If
                            ctr += 1

                            If OverdueJob.NEXT_SCHED <= DateTime.Now.AddSeconds(30) Then
                                ' Run Next Job in Overdue List
                                SetUIText_ThreadSafe(Me.StatusBar, "Running Job " + ctr.ToString + " of " + jl.Length.ToString + " - " + OverdueJob.DESCRIPTION + " Scheduled for " + OverdueJob.NEXT_SCHED.ToString())
                                Dim applicationpath As String = app.STARTUP_DIRECTORY + app.STARTUP_FILENAME
                                Dim attachmentfilepath As String = ""
                                Dim attachfilename As String = ""
                                Dim description As String = ""
                                Dim ArchiveInfo As ArchiveList()
                                strtempdirectory = "T:\HTS\REPORTS\"
                                If OverdueJob.OUTPUT_FORMAT = "XLS" Then
                                    strtempfilename = "AR" + Format(OverdueJob.JOBID, "000000") + ".XLS"
                                Else
                                    If OverdueJob.OUTPUT_FORMAT = "CSV" Then
                                        strtempfilename = "AR" + Format(OverdueJob.JOBID, "000000") + ".CSV"
                                    Else
                                        strtempfilename = "AR" + Format(OverdueJob.JOBID, "000000") + ".OUT"
                                    End If
                                End If
                                strtempfilepath = strtempdirectory & strtempfilename

                                Try
                                    reportoutcome = GenerateOutput(OverdueJob.JOBID, OverdueJob.REPORT_NAME, OverdueJob.CRITERIA, strtempdirectory, strtempfilename, OverdueJob.OUTPUT_FORMAT)

                                    If OverdueJob.OUTPUT_FORMAT = "CSV" Then
                                        System.IO.File.Copy("C:\HTS\REPORTS\OUTPUT.CSV", strtempfilepath)
                                    End If

                                    strOutputFormat = OverdueJob.OUTPUT_FORMAT

                                    attachfilename = Format(OverdueJob.JOBID, "000000") + " - " + OverdueJob.DESCRIPTION + "." + OverdueJob.OUTPUT_FORMAT
                                    attachmentfilepath = "T:\HTS\REPORTS\" & attachfilename
                                    description = "(" + OverdueJob.JOBID.ToString + ") " + OverdueJob.DESCRIPTION + " (" + OverdueJob.FREQ.ToString + OverdueJob.INTERVAL + "@" + OverdueJob.NEXT_SCHED.ToShortTimeString + ")"
                                    jobcompletion = Now.ToLocalTime
                                    archcontainer = OverdueJob.CONTAINER
                                    archfilename = Format(jobcompletion, "yyyyMMddHHmmss") & "." & strOutputFormat
                                    If reportoutcome = True Then
                                        ' Save local copy for email transmission / server upload
                                        If System.IO.File.Exists(attachmentfilepath) Then
                                            System.IO.File.Delete(attachmentfilepath)
                                        End If
                                        Thread.Sleep(1000)
                                        System.IO.File.Copy(strtempfilepath, attachmentfilepath)
                                        Thread.Sleep(1000)
                                        '  Updoad to Azure
                                        If Online.Checked Then

                                            SetUIText_ThreadSafe(Me.StatusBar, "Uploading...(" & OverdueJob.JOBID & ") " & OverdueJob.DESCRIPTION & " to Azure")
                                            uploadoutcome = UploadReportToAzure(strtempdirectory, strtempfilename, archcontainer, archfilename)
                                            ArchiveID = ws.LogArchive(OverdueJob.TYPE, archcontainer, archfilename, OverdueJob.OUTPUT_FORMAT, OverdueJob.JOBID, jobcompletion, OverdueJob.NEXT_SCHED, "", description, "Generated at " + jobcompletion.ToShortTimeString)
                                        End If
                                    End If
                                    '  Distribute to Distribution List / Update User Activity Log

                                    ArchiveInfo = ws.GetArchiveInfo(ArchiveID)
                                Catch ex As Exception
                                    reportoutcome = False
                                End Try
                            Else
                                Debug.WriteLine("Job Skipped - Not Due Yet")
                            End If
                        Next
NextApp:


                        '  Close App
                        Try
                            Dim shutdown As Boolean
                            shutdown = CloseAccessApplication(oAccess)
                            SetUIText_ThreadSafe(Me.StatusBar, "App Shutdown Successfully")
                        Catch ex As Exception
                            SetUIText_ThreadSafe(Me.StatusBar, "App Shutdown Error - Continuing")
                        End Try


                        '  Compact / Repair Database
                        '    Dim repair As Boolean
                        '    repair = CompactRepair_SecureDB(app.STARTUP_DIRECTORY + app.STARTUP_FILENAME, app.MDW_DIRECTORY + app.MDW_FILENAME, app.USERNAME, app.PASSWORD)

                    Else
                        SetUIText_ThreadSafe(Me.StatusBar, "No overdue Reports for App " = app.DESCRIPTION)
                    End If

                End If
            Next
        Catch ex As Exception
            SetUIText_ThreadSafe(Me.StatusBar, "Error during Processing: " + ex.Message.ToString())
        Finally
            '       Timer1.Enabled = True
        End Try

    End Sub

    Private Sub BGReportGenerator_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BGReportGenerator.ProgressChanged

    End Sub

    Private Sub BGReportGenerator_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BGReportGenerator.RunWorkerCompleted

        SetUIText_ThreadSafe(Me.StatusBar, "Report Generation completed...." + Now.ToString)

        If Not My.Settings.AutoShutDown And Me.WorkOffline.Checked = False Then
            My.Settings.LastCheck = Now
            Me.NextCheck.Text = GetNextCheckTime(StartTime, Offset, Freq, EndTime)
            Me.Running.Checked = False
            refreshdata()
        Else
            If My.Settings.AutoShutDown Then
                Me.Close()
                Exit Sub
                '    Application.Exit()
            End If
        End If
    End Sub

    Private Sub StartButton_Click(sender As Object, e As EventArgs) Handles StartButton.Click
        ' Enable Timer
        ' Reports will start after timer count down
        Me.WorkOffline.Checked = False
        ' Enable Buttons
        Me.StartButton.Enabled = False
        Me.StopButton.Enabled = True

    End Sub

    Private Sub StopButton_Click(sender As Object, e As EventArgs) Handles StopButton.Click
        Me.WorkOffline.Checked = True

        StopReportGenerator()
        StopReportDistributor()

        Me.StartButton.Enabled = True
        Me.StopButton.Enabled = False
    End Sub

    Private Sub RefreshButton_Click(sender As Object, e As EventArgs) Handles RefreshButton.Click
        Try
            refreshdata()
            StatusBar.Text = "Data Refreshed at " & Now.ToLocalTime.ToString
        Catch EX As Exception
            StatusBar.Text = "An Error Occured Retrieving Data from the Web Service" + EX.ToString
        End Try
    End Sub

    Private Sub RunButton_Click(sender As Object, e As EventArgs) Handles RunButton.Click
        StartReportGenerator()

    End Sub

    Private Sub SettingsButton_Click(sender As Object, e As EventArgs) Handles SettingsButton.Click
        Dim Settings As New Settings()
        Settings.ShowDialog()
        LoadSettings()
    End Sub

    Private Sub OpenCloseAccess_Click(sender As Object, e As EventArgs) Handles OpenCloseAccess.Click
        Dim oAccess As Access.Application
        Dim sDBPath As String ' Path to application mdb

        sDBPath = GetOfficeAppPath("Access.Application", "msaccess.exe")
        If sDBPath = "" Then
            MsgBox("Can't determine path to msaccess.exe", MsgBoxStyle.MsgBoxSetForeground)
        End If

        oAccess = ShellGetDB("C:\HTS\APP2K3\HTSAPP2K3.MDB", "/NOSTARTUP /WRKGRP \\SQLAPP\SHARED\HTS\HTS.MDW /USER AUTOREPORTS /PWD 4438", AppWinStyle.MinimizedFocus, 1000)

    End Sub

    Private Sub BGReportDistributor_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BGReportDistributor.DoWork

        ' Retrieve List of Jobs to Distribute
        Dim ws As New AutoReportsWCFService.ServiceClient
        Dim al() As ArchiveList
        Dim dl() As UserList
        Dim person As UserList
        Dim NumberSent As Long = 0
        Dim NumberToSend As Long = 0

        '  Retrieve List of Jobs to Distribute
       ' Dim SinceDateString = SinceDate.AddDays(-2).ToString()
        Dim SinceDateString = SinceDate.AddDays(-1).ToString()
        al = ws.GetJobsToDistribute(MaxCount, SinceDateString)
        NumberToSend = al.Length()
        '  Process Each Archive
        For Each row As ArchiveList In al


            If BGReportDistributor.CancellationPending Or NumberSent >= 15 Then
                e.Cancel = True
                Exit For
            End If

            NumberSent += 1

            ' Determine Distribution List
            If row.JOBTYPE = "ASN" Then
                dl = ws.GetASNDistributionList(row.JOBREF)
            ElseIf row.JOBTYPE = "QA" Then
                dl = ws.GetQADistributionList(row.JOBREF)
            Else
                dl = ws.GetJobDistributionList(row.JOBREF)
            End If

            ' Prepare Message to Send to Distribution List
            Dim specialmessage As String

            Dim subject As String = row.SUBJECT + " @ " + row.CREATED.ToShortDateString + " " + row.CREATED.ToShortTimeString
            Dim strfrom As String
            Dim Note As String = "PLEASE DO NOT REPLY TO SENDER"

            Dim documenturi As String = "https://htsazure.blob.core.windows.net/" + row.CONTAINER + "/" + row.FILENAME
         '   Dim websiteuri As String = "http://75.151.4.117/OnlineReports"
            Dim localfile As String = "T:\HTS\REPORTS\" + "AR" + Format(row.JOBREF, "000000")
            Dim attachfile As String = "T:\HTS\REPORTS\" + "AR" + Format(row.JOBREF, "000000") + "." + row.JOBFORMAT


            '  Create Attachment w/ useful filename
            If row.JOBFORMAT = "PDF" Then
                localfile = localfile + ".OUT"

                ' Copy localfile to create file to send
                Thread.Sleep(100)

                If System.IO.File.Exists(attachfile) Then
                    System.IO.File.Delete(attachfile)
                End If
                If System.IO.File.Exists(localfile) Then
                    If localfile <> attachfile Then
                        System.IO.File.Copy(localfile, attachfile)
                    End If
                Else
                    SetUIText_ThreadSafe(Me.StatusBar, "Local File " + localfile + " Not found ")
                End If
                Thread.Sleep(100)

            Else
                localfile = localfile + ".XLS"
            End If
            strfrom = My.Settings.FromUserName
            '   strfrom = "autoreports@ssitroy.com"

            specialmessage = "<p>IMPORTANT NOTE</p>"
            If Note.Length > 0 Then
                specialmessage = specialmessage & Note & Chr(10) & Chr(13)
            End If

            specialmessage = specialmessage & "<p><a href=" & documenturi & ">Click Here to View Report</a></p>" & Chr(10) & Chr(13)

            specialmessage = specialmessage & "<p>You are receiving this automated email from Heat Treating Services.</p>" & Chr(10) & Chr(13)
            specialmessage = specialmessage & "<p>To ensure proper delivery, please be sure to add autoreports@htsmi.com to your trusted senders list.</p>" & Chr(10) & Chr(13)
            specialmessage = specialmessage & "<p>For any production or scheduling questions, please contact your plant representative or call (248) 858-2230.</p>" & Chr(10) & Chr(13)
            specialmessage = specialmessage & "<p>Thank you</p>"

            ' specialmessage = specialmessage + "<p>Click here to Login to <a href=" + websiteuri + "> HTS Online</a></p>" + Chr(10) & Chr(13)

            ' Update Activity Log / Populate Recipients
            Dim distlist As String = ""

            specialmessage = specialmessage & "<p>Current Distribution List: (Select a Name Below to send a message to a recipient)</p>"
            specialmessage = specialmessage & "<ul>"
            For Each person In dl
                If distlist <> "" Then
                    distlist = distlist + ","
                End If
                '  Do Not populate distlist if it exceeds storage max.  email will still be sent but just not included in response string in email
                If distlist.Length + person.EMAIL.Length < 255 Then
                    distlist = distlist + person.EMAIL
                End If
                specialmessage = specialmessage & "<li>" & person.FULLNAME & " (" & person.EMAIL & ")"
                ws.LogActivity(row.JOBTYPE, row.ARCHIVEID, person.USERID, row.JOBREF, "Email Sent at " + Now.ToShortTimeString(), DateTime.Now)
            Next
            specialmessage = specialmessage & "</ul>"

            SetUIText_ThreadSafe(Me.StatusBar, "Emailing..." + NumberSent.ToString + " of " + NumberToSend.ToString + " " + row.JOBREF.ToString + " " + row.SUBJECT & " to " + distlist)

            specialmessage = specialmessage & "<p>Click here to send a message to <a href=mailto:" & distlist & "?subject=Reply%20Message%20to%20AutoReport>" & "All Recipients</a></p>" & Chr(10) & Chr(13)

            '     specialmessage = specialmessage & "<p>Current Distribution List => " & distlist & "</p>" & Chr(10) & Chr(13)

            specialmessage = specialmessage & "<p>To Add or Change Distribution, Click Here to send a message to <a href=mailto:autoreports@ssitroy.com;fshepard@htsmi.com;jwhaley@htsmi.com" + "?subject=Message%20to%20Admins%20Message%20to%20AutoReport>" + " System Administrators</a></p>" & Chr(10) & Chr(13) & Chr(10) & Chr(13)

            specialmessage = specialmessage & "<p>Message Sent from HTSAUTO at " & DateTime.Now.ToString & " / Ref: " & row.ARCHIVEID & "</p>" & Chr(10) & Chr(13)
            ' Send Email


            Dim result As Boolean = SendEmailX(dl, strfrom, subject, specialmessage, attachfile)
            '  This is commented out during testing
            If "a" = "a" Then

                ' Mark Archive as Sent

                ws.MarkArchiveAsSent(row.ARCHIVEID, distlist)
                If row.JOBTYPE = "ASN" Or row.JOBTYPE = "QA" Then
                    ws.MarkASNAsSent(row.JOBREF, DateTime.Now.ToShortTimeString, row.CONTAINER, row.FILENAME, distlist)
                End If
            End If
        Next




    End Sub

    Private Sub BGReportDistributor_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BGReportDistributor.RunWorkerCompleted

        SetUIText_ThreadSafe(Me.StatusBar, "Distribution Completed at ..." + Now.ToString)

        If Me.WorkOffline.Checked = False And Not My.Settings.AutoShutDown Then
            My.Settings.LastDistributed = Now
            Me.NextDistribute.Text = GetNextCheckTime(StartTime, DistributeOffset, DistributeFreq, EndTime)
            Me.Distributing.Checked = False
            refreshdata()
        Else
            If My.Settings.AutoShutDown Then
                Me.Close()
                '  Exit()
                '  Application.Exit()
            End If
        End If
    End Sub

    Private Sub DistributeButton_Click(sender As Object, e As EventArgs) Handles DistributeButton.Click
        StartReportDistributor()
    End Sub

    Private Sub BGMessenger_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BGMessenger.DoWork

        ' Retrieve List of Messages To Send Out
        Dim ws As New AutoReportsWCFService.ServiceClient
        Dim ml() As MessageList
        Dim dl() As UserList
        Dim person As UserList
        Dim NumberSent As Long = 0
        Dim NumberToSend As Long = 0
        Dim distlist As String
        '  Retrieve List of Jobs to Distribute
        Dim SinceDateString As String = DateTime.Now.AddDays(-2).ToString()

        Try
            
            ml = ws.GetOverdueMessages(MaxCount, SinceDateString)
            NumberToSend = ml.Length()
            '  Process Each Archive
            For Each row As MessageList In ml


                If BGMessenger.CancellationPending Or NumberSent >= 15 Then
                    e.Cancel = True
                    Exit For
                End If

                NumberSent += 1

                ' Determine Distribution List
                
                distlist = row.DISTLIST

                  '  Need to split ccDistList by ";"
                Dim dlLIST As STRING() 
                dlLIST = distlist.Split(";")
               
             '  Need to split ccDistList by ";"
          
                SetUIText_ThreadSafe(Me.StatusBar, "Sending Messages..." + NumberSent.ToString + " of " + NumberToSend.ToString + " " + row.ID.ToString + " " + row.SUBJECT & " to " + distlist)

                ' Prepare Message to Send to Distribution List
                Dim specialmessage As String

                Dim subject As String = "(" + row.ACCESS_USERID + ") " + row.SUBJECT + " @ " + row.CREATED.ToShortDateString + " " + row.CREATED.ToShortTimeString
                Dim Note As String = "DO NOT REPLY"


                specialmessage = "<p>IMPORTANT NOTE</p>"
                If Note.Length > 0 Then
                    specialmessage = specialmessage & Note & "<br>"
                End If
                Dim BODY As STRING
                BODY = Replace(row.BODY,vbCrLf,"<BR>")
             
                specialmessage = specialmessage + BODY
                specialmessage = specialmessage & "<p>You are receiving this message via email from Heat Treating Services.  If you would like to be removed from distribution or would like someone else to be added, please contact Franklin Shepard or simply reply to this email.</p>" & Chr(10) & Chr(13)
                specialmessage = specialmessage & "<p>To ensure proper delivery, please be sure to add autoreports@ssitroy.com to your trusted senders list.</p>" & Chr(10) & Chr(13)
                specialmessage = specialmessage & "<p>For any production or scheduling questions, please contact your plant representative or call (248) 858-2230.</p>" & Chr(10) & Chr(13)
                specialmessage = specialmessage & "<p>Thank you</p>"


                specialmessage = specialmessage & "<p>Click here to send a message to <a href=mailto:" & distlist & "?subject=Reply%20Message%20to%20AutoReport>" & "All Recipients</a></p>" & Chr(10) & Chr(13)

                specialmessage = specialmessage & "<p>Click here to send a message to <a href=mailto:autoreports@ssitroy.com;fshepard@htsmi.com" + "?subject=Message%20to%20Admins%20Message%20to%20AutoReport>" + "System Administrators</a></p>" & Chr(10) & Chr(13)

                specialmessage = specialmessage & "<p>Message Sent from HTSAUTO at " & DateTime.Now.ToString & " / Ref: " & row.ID & "</p>" & Chr(10) & Chr(13)


                ' specialmessage = specialmessage + "<p>Click here to Login to <a href=" + websiteuri + "> HTS Online</a></p>" + Chr(10) & Chr(13)
                Dim emailresult As Boolean
                Dim sentresult As Boolean

                
            '    emailresult = SendEmailToGroup(row.DISTLIST, "autoreports@htsmi.com", subject, row.BODY,"")
                emailresult = SendEmailToGroup(row.DISTLIST,"autoreports@htsmi.com", subject, specialmessage,"")
           '     emailresult = ws.SendMessage(row.DISTLIST, subject, row.BODY)
                If emailresult = True Then
                    sentresult = ws.MarkQAMessageAsSent(row.ID)
                    SetUIText_ThreadSafe(Me.StatusBar, "Message Sent")
                Else
                    SetUIText_ThreadSafe(Me.StatusBar, "Error during Send")
                End If

            Next

        Catch ex As Exception
            SetUIText_ThreadSafe(Me.StatusBar, "Error during Message Distribution: " + ex.Message.ToString())
        Finally
            Timer1.Enabled = True
        End Try


    End Sub

    Private Sub Messenger_Click(sender As Object, e As EventArgs) Handles Messenger.Click
        StartMessenger()
    End Sub

    Private Sub BGMessenger_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BGMessenger.RunWorkerCompleted
        SetUIText_ThreadSafe(Me.StatusBar, "Messenger Service Completed at ..." + Now.ToString)

        If Me.WorkOffline.Checked = False And Not My.Settings.AutoShutDown Then
            Me.SendingMessages.Checked = False
            My.Settings.LastMessageRun = Now
            Me.NextMessageRun.Text = GetNextCheckTime(StartTime, DistributeOffset - 1, DistributeFreq, EndTime)
            Me.SendingMessages.Checked = False
            refreshdata()
        Else
            If My.Settings.AutoShutDown Then
                Me.Close()
                Exit Sub
                '                Application.Exit()
            End If
        End If
    End Sub

    Private Sub SendingMessages_CheckedChanged(sender As Object, e As EventArgs) Handles SendingMessages.CheckedChanged

    End Sub

    Private Sub MessageListBindingSource_CurrentChanged(sender As Object, e As EventArgs) Handles MessageListBindingSource.CurrentChanged

    End Sub


End Class
