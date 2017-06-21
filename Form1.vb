Imports Microsoft.Office
Imports Microsoft.Office.interop
Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

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
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton3 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton4 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton5 As System.Windows.Forms.RadioButton
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Timer1 As System.Timers.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton
        Me.RadioButton2 = New System.Windows.Forms.RadioButton
        Me.RadioButton3 = New System.Windows.Forms.RadioButton
        Me.RadioButton4 = New System.Windows.Forms.RadioButton
        Me.RadioButton5 = New System.Windows.Forms.RadioButton
        Me.Button1 = New System.Windows.Forms.Button
        Me.Timer1 = New System.Timers.Timer
        CType(Me.Timer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'RadioButton1
        '
        Me.RadioButton1.Location = New System.Drawing.Point(96, 32)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(150, 24)
        Me.RadioButton1.TabIndex = 0
        Me.RadioButton1.Text = "RadioButton1"
        '
        'RadioButton2
        '
        Me.RadioButton2.Location = New System.Drawing.Point(96, 64)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(150, 24)
        Me.RadioButton2.TabIndex = 1
        Me.RadioButton2.Text = "RadioButton2"
        '
        'RadioButton3
        '
        Me.RadioButton3.Location = New System.Drawing.Point(96, 96)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(150, 24)
        Me.RadioButton3.TabIndex = 2
        Me.RadioButton3.Text = "RadioButton3"
        '
        'RadioButton4
        '
        Me.RadioButton4.Location = New System.Drawing.Point(96, 128)
        Me.RadioButton4.Name = "RadioButton4"
        Me.RadioButton4.Size = New System.Drawing.Size(150, 24)
        Me.RadioButton4.TabIndex = 3
        Me.RadioButton4.Text = "RadioButton4"
        '
        'RadioButton5
        '
        Me.RadioButton5.Location = New System.Drawing.Point(96, 160)
        Me.RadioButton5.Name = "RadioButton5"
        Me.RadioButton5.Size = New System.Drawing.Size(150, 24)
        Me.RadioButton5.TabIndex = 4
        Me.RadioButton5.Text = "RadioButton5"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(88, 200)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(112, 23)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "Button1"
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 10000
        Me.Timer1.SynchronizingObject = Me
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(624, 258)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.RadioButton5)
        Me.Controls.Add(Me.RadioButton4)
        Me.Controls.Add(Me.RadioButton3)
        Me.Controls.Add(Me.RadioButton2)
        Me.Controls.Add(Me.RadioButton1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        CType(Me.Timer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private m_sAction As String

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles MyBase.Load
        RadioButton1.Text = "Print report"
        RadioButton2.Text = "Preview report"
        RadioButton3.Text = "Show form"
        RadioButton4.Text = "Print report (Security)"
        RadioButton5.Text = "Preview report (Runtime)"
        Button1.Text = "Go!"
        m_sAction = "Print report"
    End Sub

    Private Sub RadioButtons_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles RadioButton1.Click, RadioButton2.Click, RadioButton3.Click, _
        RadioButton4.Click, RadioButton5.Click
        m_sAction = sender.Text 'Store the text for the selected radio button
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles Button1.Click
        ' Calls the associated procedure to automate Access, based
        ' on the selected radio button on the form.
        Select Case m_sAction
            Case "Print report" : Print_Report()
            Case "Preview report" : Preview_Report()
            Case "Show form" : Show_Form()
            Case "Print report (Security)" : Print_Report_Security()
            Case "Preview report (Runtime)" : Preview_Report_Runtime()
        End Select
    End Sub

    Public Function ShellGetDB(ByVal sDBPath As String, _
        Optional ByVal sCmdLine As String = vbNullString, _
        Optional ByVal enuWindowStyle As Microsoft.VisualBasic.AppWinStyle _
            = AppWinStyle.MinimizedFocus, _
        Optional ByVal iSleepTime As Integer = 1000) As Access.Application

        'Launches a new instance of Access with a database (sDBPath)
        'using the Shell function then returns the Application object
        'via calling: GetObject(sDBPath). Returns the Application
        'object of the new instance of Access, assuming that sDBPath
        'is not already opened in another instance of Access. To ensure
        'the Application object of the new instance is returned, make
        'sure sDBPath is not already opened in another instance of Access.
        '
        'Example:
        'Dim oAccess As Access.Application
        'oAccess = ShellGetDB("c:\mydb.mdb")

        ' Enable an error handler for this procedure:
        On Error GoTo ErrorHandler

        Dim oAccess As Access.Application
        Dim sAccPath As String 'path to msaccess.exe

        ' Obtain the path to msaccess.exe:
        sAccPath = GetOfficeAppPath("Access.Application", "msaccess.exe")
        If sAccPath = "" Then
            MsgBox("Can't determine path to msaccess.exe", _
                MsgBoxStyle.MsgBoxSetForeground)
            Return Nothing
        End If

        ' Make sure specified database (sDBPath) exists:
        If Not System.IO.File.Exists(sDBPath) Then
            MsgBox("Can't find the file '" & sDBPath & "'", _
                MsgBoxStyle.MsgBoxSetForeground)
            Return Nothing
        End If

        ' Start a new instance of Access using sDBPath and sCmdLine:
        If sCmdLine = vbNullString Then
            sCmdLine = Chr(34) & sDBPath & Chr(34)
        Else
            sCmdLine = Chr(34) & sDBPath & Chr(34) & " " & sCmdLine
        End If
        Shell(Pathname:=sAccPath & " " & sCmdLine, _
            Style:=enuWindowStyle)
        'Note: It is advised that the Style argument of the Shell
        'function be used to give focus to Access.

        ' Move focus back to this form. This ensures that Access
        ' registers itself in the ROT, allowing GetObject to find it:
        AppActivate(Title:=Me.Text)

        ' Pause to allow database to open:
        System.Threading.Thread.Sleep(iSleepTime)

        ' Obtain Application object of the instance of Access
        ' that has the database open:
        oAccess = GetObject(sDBPath)

        Return oAccess
ErrorCleanup:
        ' Try to quit Access due to an unexpected error:
        On Error Resume Next
        oAccess.Quit(Option:=Access.AcQuitOption.acQuitSaveNone)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oAccess)
        oAccess = Nothing

        Return Nothing
ErrorHandler:
        MsgBox(Err.Number & ": " & Err.Description, _
            MsgBoxStyle.MsgBoxSetForeground, "Error Handler")
        Resume ErrorCleanup
    End Function

    Public Function ShellGetApp(Optional ByVal sCmdLine As String = vbNullString, _
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
            MsgBox("Can't determine path to msaccess.exe", _
                MsgBoxStyle.MsgBoxSetForeground)
            Return Nothing
        End If

        ' Start a new instance of Access using sCmdLine:
        If sCmdLine = vbNullString Then
            sCmdLine = sAccPath
        Else
            sCmdLine = sAccPath & " " & sCmdLine
        End If
        Shell(Pathname:=sCmdLine, Style:=enuWindowStyle)
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
                MsgBox("GetObject failed. Process ended.", _
                    MsgBoxStyle.MsgBoxSetForeground)
            End If
        Else 'iSection = 0 so use normal error handling:
            MsgBox(Err.Number & ": " & Err.Description, _
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
        On Error GoTo ErrorHandler

        Dim oReg As Microsoft.Win32.RegistryKey = _
            Microsoft.Win32.Registry.LocalMachine
        Dim oKey As Microsoft.Win32.RegistryKey
        Dim sCLSID As String
        Dim sPath As String
        Dim iPos As Integer

        ' First, get the clsid from the progid from the registry key
        ' HKEY_LOCAL_MACHINE\Software\Classes\<PROGID>\CLSID:
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
ErrorHandler:
        Return ""
    End Function

    Private Sub Print_Report()
        'Prints the "Summary of Sales by Year" report in Northwind.mdb.

        ' Enable an error handler for this procedure:
        On Error GoTo ErrorHandler

        Dim oAccess As Access.Application
        Dim sDBPath As String 'path to Northwind.mdb
        Dim sReport As String 'name of report to print

        sReport = "Summary of Sales by Year"

        ' Start a new instance of Access for automation:
        oAccess = New Access.ApplicationClass

        ' Determine the path to Northwind.mdb:
        sDBPath = oAccess.SysCmd(Action:=Access.AcSysCmdAction.acSysCmdAccessDir)
        sDBPath = sDBPath & "Samples\Northwind.mdb"

        ' Open Northwind.mdb in shared mode:
        oAccess.OpenCurrentDatabase(filepath:=sDBPath, Exclusive:=False)

        ' Select the report name in the database window and give focus
        ' to the database window:
        oAccess.DoCmd.SelectObject(ObjectType:=Access.AcObjectType.acReport, _
            ObjectName:=sReport, InDatabaseWindow:=True)

        ' Print the report:
        oAccess.DoCmd.OpenReport(ReportName:=sReport, _
            View:=Access.AcView.acViewNormal)

Cleanup:
        ' Quit Access and release object:
        On Error Resume Next
        oAccess.Quit(Option:=Access.AcQuitOption.acQuitSaveNone)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oAccess)
        oAccess = Nothing

        Exit Sub
ErrorHandler:
        MsgBox(Err.Number & ": " & Err.Description, _
            MsgBoxStyle.MsgBoxSetForeground, "Error Handler")
        ' Try to quit Access due to an unexpected error:
        Resume Cleanup
    End Sub

    Private Sub Preview_Report()
        'Previews the "Summary of Sales by Year" report in Northwind.mdb.

        ' Enable an error handler for this procedure:
        On Error GoTo ErrorHandler

        Dim oAccess As Access.Application
        Dim oForm As Access.AccessObject
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
            oAccess.DoCmd.Close(ObjectType:=Access.AcObjectType.acForm, _
                ObjectName:=oForm.Name, _
                Save:=Access.AcCloseSave.acSaveNo)
        Next
        If Not oForm Is Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
        End If
        oForm = Nothing

        ' Select the report name in the database window and give focus
        ' to the database window:
        oAccess.DoCmd.SelectObject(ObjectType:=Access.AcObjectType.acReport, _
            ObjectName:=sReport, InDatabaseWindow:=True)

        ' Maximize the Access window:
        oAccess.RunCommand(Command:=Access.AcCommand.acCmdAppMaximize)

        ' Preview the report:
        oAccess.DoCmd.OpenReport(ReportName:=sReport, _
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

        '        oAccess.DoCmd.OutputTo(Access.AcOutputObjectType.acOutputReport, sReport, System.Windows.Forms.DataFormats.Rtf, "c:\myreport.rtf", True)
        oAccess.DoCmd.OutputTo(Access.AcOutputObjectType.acOutputReport, sReport, System.Windows.Forms.DataFormats.Rtf, ",c:\myreport.rft", True, , , Access.AcExportQuality.acExportQualityScreen)

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
        MsgBox(Err.Number & ": " & Err.Description, _
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
            oAccess.DoCmd.Close(ObjectType:=Access.AcObjectType.acForm, _
                ObjectName:=oForm.Name, _
                Save:=Access.AcCloseSave.acSaveNo)
        Next
        If Not oForm Is Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
        End If
        oForm = Nothing

        ' Select the form name in the database window and give focus
        ' to the database window:
        oAccess.DoCmd.SelectObject(ObjectType:=Access.AcObjectType.acForm, _
            ObjectName:=sForm, InDatabaseWindow:=True)

        ' Show the form:
        oAccess.DoCmd.OpenForm(FormName:=sForm, _
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
        oAccess.DoCmd.SelectObject(ObjectType:=Access.AcObjectType.acForm, _
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
        MsgBox(Err.Number & ": " & Err.Description, _
            MsgBoxStyle.MsgBoxSetForeground, "Error Handler")
        Resume ErrorCleanup
    End Sub

    Private Sub Print_Report_Security()
        'Shows how to automate Access when user-level
        'security is enabled and you wish to avoid the Logon
        'dialog asking for user name and password. In this 
        'example we're assuming default security so we simply
        'pass the Admin user with a blank password to print the 
        '"Summary of Sales by Year" report in Northwind.mdb.

        ' Enable an error handler for this procedure:
        On Error GoTo ErrorHandler

        Dim oAccess As Access.Application
        Dim sDBPath As String 'path to Northwind.mdb
        Dim sUser As String 'user name for Access security
        Dim sPwd As String 'user password for Access security
        Dim sReport As String 'name of report to print

        sReport = "Summary of Sales by Year"

        ' Determine the path to Northwind.mdb:
        sDBPath = GetOfficeAppPath("Access.Application", "msaccess.exe")
        If sDBPath = "" Then
            MsgBox("Can't determine path to msaccess.exe", _
                MsgBoxStyle.MsgBoxSetForeground)
            Exit Sub
        End If
        sDBPath = Microsoft.VisualBasic.Left(sDBPath, _
            Len(sDBPath) - Len("msaccess.exe")) & "Samples\Northwind.mdb"
        If Not System.IO.File.Exists(sDBPath) Then
            MsgBox("Can't find the file '" & sDBPath & "'", _
                MsgBoxStyle.MsgBoxSetForeground)
            Exit Sub
        End If

        ' Specify the user name and password for the Access workgroup
        ' information file, which is used to implement Access user-level security.
        ' The file by default is named System.mdw and can be specified
        ' using the /wrkgrp command-line switch. This example assumes
        ' default security and therefore does not specify a workgroup
        ' information file and uses Admin with no password:
        sUser = "user_name"
        sPwd = "my_password"

        ' Start a new instance of Access with user name and password:
        oAccess = ShellGetDB(sDBPath, "/user " & sUser & " /pwd " & sPwd)
        'or
        'oAccess = ShellGetApp(Chr(34) & sDBPath & Chr(34) & " /user " & sUser & " /pwd " & sPwd)

        ' Select the report name in the database window and give focus
        ' to the database window:
        oAccess.DoCmd.SelectObject(ObjectType:=Access.AcObjectType.acReport, _
            ObjectName:=sReport, InDatabaseWindow:=True)

        ' Print the report:
        oAccess.DoCmd.OpenReport(ReportName:=sReport, _
            View:=Access.AcView.acViewNormal)

Cleanup:
        ' Quit Access and release object:
        On Error Resume Next
        oAccess.Quit(Option:=Access.AcQuitOption.acQuitSaveNone)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oAccess)
        oAccess = Nothing

        Exit Sub
ErrorHandler:
        MsgBox(Err.Number & ": " & Err.Description, _
            MsgBoxStyle.MsgBoxSetForeground, "Error Handler")
        ' Try to quit Access due to an unexpected error:
        Resume Cleanup
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
            MsgBox("Can't determine path to msaccess.exe", _
                MsgBoxStyle.MsgBoxSetForeground)
            Exit Sub
        End If
        sDBPath = Microsoft.VisualBasic.Left(sDBPath, _
            Len(sDBPath) - Len("msaccess.exe")) & "Samples\Northwind.mdb"
        If Not System.IO.File.Exists(sDBPath) Then
            MsgBox("Can't find the file '" & sDBPath & "'", _
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
            oAccess.DoCmd.Close(ObjectType:=Access.AcObjectType.acForm, _
                ObjectName:=oForm.Name, _
                Save:=Access.AcCloseSave.acSaveNo)
        Next
        If Not oForm Is Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
        End If
        oForm = Nothing

        ' Select the report name in the database window and give focus
        ' to the database window:
        oAccess.DoCmd.SelectObject(ObjectType:=Access.AcObjectType.acReport, _
            ObjectName:=sReport, InDatabaseWindow:=True)

        ' Maximize the Access window:
        oAccess.RunCommand(Command:=Access.AcCommand.acCmdAppMaximize)

        ' Preview the report:
        oAccess.DoCmd.OpenReport(ReportName:=sReport, _
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
        MsgBox(Err.Number & ": " & Err.Description, _
            MsgBoxStyle.MsgBoxSetForeground, "Error Handler")
        Resume ErrorCleanup
    End Sub


    Private Sub Timer1_Elapsed(ByVal sender As System.Object, ByVal e As System.Timers.ElapsedEventArgs) Handles Timer1.Elapsed
        Button1_Click(sender, e)
    End Sub
End Class
