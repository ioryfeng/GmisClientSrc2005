Public Class FConsole
    Inherits System.Windows.Forms.Form

#Region " Windows 窗体设计器生成的代码 "

    Public Sub New()
        MyBase.New()

        '该调用是 Windows 窗体设计器所必需的。
        InitializeComponent()

        '在 InitializeComponent() 调用之后添加任何初始化

    End Sub

    '窗体重写处置以清理组件列表。
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意：以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改此过程。
    '不要使用代码编辑器修改它。
    Friend WithEvents imlButton As System.Windows.Forms.ImageList
    Friend WithEvents ImlButton1 As System.Windows.Forms.ImageList
    Friend WithEvents btnLogin As System.Windows.Forms.Label
    Friend WithEvents btnMail As System.Windows.Forms.Label
    Friend WithEvents btnExit As System.Windows.Forms.Label
    Friend WithEvents btnQuery As System.Windows.Forms.Label
    Friend WithEvents btnWorkLog As System.Windows.Forms.Label
    Friend WithEvents btnSystem As System.Windows.Forms.Label
    Friend WithEvents btnDocument As System.Windows.Forms.Label
    Friend WithEvents btnBusiness As System.Windows.Forms.Label
    Friend WithEvents tmrMessage As System.Windows.Forms.Timer
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents btnRefresh As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FConsole))
        Me.imlButton = New System.Windows.Forms.ImageList(Me.components)
        Me.btnLogin = New System.Windows.Forms.Label
        Me.btnMail = New System.Windows.Forms.Label
        Me.btnExit = New System.Windows.Forms.Label
        Me.ImlButton1 = New System.Windows.Forms.ImageList(Me.components)
        Me.btnQuery = New System.Windows.Forms.Label
        Me.btnWorkLog = New System.Windows.Forms.Label
        Me.btnSystem = New System.Windows.Forms.Label
        Me.btnDocument = New System.Windows.Forms.Label
        Me.btnBusiness = New System.Windows.Forms.Label
        Me.tmrMessage = New System.Windows.Forms.Timer(Me.components)
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.btnRefresh = New System.Windows.Forms.Label
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.SuspendLayout()
        '
        'imlButton
        '
        Me.imlButton.ImageStream = CType(resources.GetObject("imlButton.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imlButton.TransparentColor = System.Drawing.Color.Transparent
        Me.imlButton.Images.SetKeyName(0, "")
        Me.imlButton.Images.SetKeyName(1, "")
        Me.imlButton.Images.SetKeyName(2, "")
        Me.imlButton.Images.SetKeyName(3, "")
        Me.imlButton.Images.SetKeyName(4, "")
        Me.imlButton.Images.SetKeyName(5, "")
        '
        'btnLogin
        '
        Me.btnLogin.ImageIndex = 0
        Me.btnLogin.ImageList = Me.imlButton
        Me.btnLogin.Location = New System.Drawing.Point(16, 168)
        Me.btnLogin.Name = "btnLogin"
        Me.btnLogin.Size = New System.Drawing.Size(114, 36)
        Me.btnLogin.TabIndex = 0
        '
        'btnMail
        '
        Me.btnMail.ImageIndex = 4
        Me.btnMail.ImageList = Me.imlButton
        Me.btnMail.Location = New System.Drawing.Point(16, 208)
        Me.btnMail.Name = "btnMail"
        Me.btnMail.Size = New System.Drawing.Size(114, 36)
        Me.btnMail.TabIndex = 1
        '
        'btnExit
        '
        Me.btnExit.ImageIndex = 2
        Me.btnExit.ImageList = Me.imlButton
        Me.btnExit.Location = New System.Drawing.Point(16, 248)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(114, 36)
        Me.btnExit.TabIndex = 2
        '
        'ImlButton1
        '
        Me.ImlButton1.ImageStream = CType(resources.GetObject("ImlButton1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImlButton1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImlButton1.Images.SetKeyName(0, "")
        Me.ImlButton1.Images.SetKeyName(1, "")
        Me.ImlButton1.Images.SetKeyName(2, "")
        Me.ImlButton1.Images.SetKeyName(3, "")
        Me.ImlButton1.Images.SetKeyName(4, "")
        Me.ImlButton1.Images.SetKeyName(5, "")
        Me.ImlButton1.Images.SetKeyName(6, "")
        Me.ImlButton1.Images.SetKeyName(7, "")
        Me.ImlButton1.Images.SetKeyName(8, "")
        Me.ImlButton1.Images.SetKeyName(9, "")
        '
        'btnQuery
        '
        Me.btnQuery.ImageIndex = 0
        Me.btnQuery.ImageList = Me.ImlButton1
        Me.btnQuery.Location = New System.Drawing.Point(496, 136)
        Me.btnQuery.Name = "btnQuery"
        Me.btnQuery.Size = New System.Drawing.Size(123, 58)
        Me.btnQuery.TabIndex = 3
        '
        'btnWorkLog
        '
        Me.btnWorkLog.ImageIndex = 2
        Me.btnWorkLog.ImageList = Me.ImlButton1
        Me.btnWorkLog.Location = New System.Drawing.Point(576, 192)
        Me.btnWorkLog.Name = "btnWorkLog"
        Me.btnWorkLog.Size = New System.Drawing.Size(123, 58)
        Me.btnWorkLog.TabIndex = 4
        '
        'btnSystem
        '
        Me.btnSystem.ImageIndex = 4
        Me.btnSystem.ImageList = Me.ImlButton1
        Me.btnSystem.Location = New System.Drawing.Point(568, 272)
        Me.btnSystem.Name = "btnSystem"
        Me.btnSystem.Size = New System.Drawing.Size(123, 58)
        Me.btnSystem.TabIndex = 5
        '
        'btnDocument
        '
        Me.btnDocument.ImageIndex = 6
        Me.btnDocument.ImageList = Me.ImlButton1
        Me.btnDocument.Location = New System.Drawing.Point(496, 336)
        Me.btnDocument.Name = "btnDocument"
        Me.btnDocument.Size = New System.Drawing.Size(123, 58)
        Me.btnDocument.TabIndex = 6
        '
        'btnBusiness
        '
        Me.btnBusiness.ImageIndex = 8
        Me.btnBusiness.ImageList = Me.ImlButton1
        Me.btnBusiness.Location = New System.Drawing.Point(368, 368)
        Me.btnBusiness.Name = "btnBusiness"
        Me.btnBusiness.Size = New System.Drawing.Size(123, 58)
        Me.btnBusiness.TabIndex = 7
        '
        'tmrMessage
        '
        Me.tmrMessage.Interval = 300000
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "")
        Me.ImageList1.Images.SetKeyName(1, "")
        '
        'btnRefresh
        '
        Me.btnRefresh.ImageIndex = 0
        Me.btnRefresh.ImageList = Me.ImageList1
        Me.btnRefresh.Location = New System.Drawing.Point(8, 424)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(66, 49)
        Me.btnRefresh.TabIndex = 9
        Me.btnRefresh.Visible = False
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.SystemColors.Desktop
        Me.Panel3.Location = New System.Drawing.Point(0, 648)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(792, 8)
        Me.Panel3.TabIndex = 13
        '
        'FConsole
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.ClientSize = New System.Drawing.Size(800, 479)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.btnRefresh)
        Me.Controls.Add(Me.btnBusiness)
        Me.Controls.Add(Me.btnDocument)
        Me.Controls.Add(Me.btnSystem)
        Me.Controls.Add(Me.btnWorkLog)
        Me.Controls.Add(Me.btnQuery)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnMail)
        Me.Controls.Add(Me.btnLogin)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "FConsole"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "主控制台 - 担保业务管理软件"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private _mainForm As frmMainInterface
    Private Sub Button_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.MouseEnter, btnLogin.MouseEnter, btnMail.MouseEnter
        Me.Cursor = Cursors.Hand

        If CType(sender, Label).ImageIndex < imlButton.Images.Count() - 1 Then
            CType(sender, Label).ImageIndex += 1
        End If
    End Sub

    Private Sub Button_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.MouseLeave, btnLogin.MouseLeave, btnMail.MouseLeave
        If CType(sender, Label).ImageIndex > 0 Then
            CType(sender, Label).ImageIndex -= 1
        End If
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub Button2_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRefresh.MouseEnter
        If CType(sender, Label).ImageIndex < Me.ImageList1.Images.Count() - 1 Then
            CType(sender, Label).ImageIndex += 1
        End If
        Me.Cursor = Cursors.Hand
    End Sub

    Private Sub Button2_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRefresh.MouseLeave
        If CType(sender, Label).ImageIndex > 0 Then
            CType(sender, Label).ImageIndex -= 1
        End If
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub Button1_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBusiness.MouseEnter, btnDocument.MouseEnter, btnQuery.MouseEnter, btnSystem.MouseEnter, btnWorkLog.MouseEnter
        If CType(sender, Label).ImageIndex < ImlButton1.Images.Count() - 1 Then
            CType(sender, Label).ImageIndex += 1
        End If
        Me.Cursor = Cursors.Hand
    End Sub

    Private Sub Button1_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBusiness.MouseLeave, btnDocument.MouseLeave, btnQuery.MouseLeave, btnSystem.MouseLeave, btnWorkLog.MouseLeave
        If CType(sender, Label).ImageIndex > 0 Then
            CType(sender, Label).ImageIndex -= 1
        End If
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnMail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMail.Click
        Try
            System.Diagnostics.Process.Start("OUTLOOK.EXE")
            Me.WindowState = FormWindowState.Minimized

        Catch ex As Exception
            MessageBox.Show("无法打开 OutLook，请确定是否安装。", "打开邮箱失败", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try
    End Sub

    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click
        Dim login As SWDialogBox.LoginBox = New SWDialogBox.LoginBox()

        Try
            '读取配置
            login.UserName = SWConsole.Configuration.GetAppValue("LoginBox.UserName")
            login.Server = SWConsole.Configuration.GetAppValue("LoginBox.Server")
        Catch
        End Try

        AddHandler login.Logined, AddressOf LoginedProcess

        If login.Login() Then
            If (Not _mainForm Is Nothing) AndAlso (Not _mainForm.IsDisposed) Then
                _mainForm.Close()
            End If

            Me.Refresh()
        End If
    End Sub

    Private Sub ShowMain(ByVal status As String)
        Try
            Me.Cursor = Cursors.AppStarting

            If _mainForm Is Nothing OrElse _mainForm.IsDisposed Then
                _mainForm = New frmMainInterface()
            End If

            setFormMenu(_mainForm)
            _mainForm.SetStatus(status)
            _mainForm.Show()
            _mainForm.Activate()
            Me.WindowState = FormWindowState.Minimized
        Catch ex As Exception
            SWDialogBox.MessageBox.Show("*999", ex.Source, ex.Message, ex.StackTrace)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btnWorkLog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWorkLog.Click
        Me.ShowMain("WorkLog")
        Me.WindowState = FormWindowState.Minimized

    End Sub

    Private Sub btnQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuery.Click
        'Me.ShowMain("Query")
        'Me.WindowState = FormWindowState.Minimized

        Me.WindowState = FormWindowState.Minimized

        Try
            Me.Cursor = Cursors.AppStarting

            Dim strComnd As String
            strComnd = "GMISBI.exe """ & UserName & """ """ & strServer & """"
            Shell(strComnd, AppWinStyle.NormalFocus)
            'Shell("""SWUserManager.exe"" " & UserName & """ """ & strPasswd & """ " & strServer, AppWinStyle.Hide)
        Catch ex As Exception
            SWDialogBox.MessageBox.Show("*999", ex.Source, ex.Message, ex.StackTrace)
        Finally
            Me.Cursor = Cursors.Default
        End Try

    End Sub

    Private Sub btnDocument_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDocument.Click
        Me.ShowMain("Document")
        Me.WindowState = FormWindowState.Minimized

    End Sub

    Private Sub btnSystem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSystem.Click
        Me.ShowMain("System")
        Me.WindowState = FormWindowState.Minimized

    End Sub

    Private Sub btnBusiness_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBusiness.Click
        Me.ShowMain("Business")
        Me.WindowState = FormWindowState.Minimized

    End Sub

    Public Shadows Sub Refresh()
        MyBase.Refresh()

        Try

            ''2007-10-16 yjf add 扫描预警消息
            'Try
            '    gWs.ScanTimingTask()
            'Catch
            'End Try

            'Dim dsTmp As Data.DataSet
            'Dim iCounter, i As Integer

            'Me.Cursor = Cursors.AppStarting

            'Me.lvwMessage.Items.Clear()

            'dsTmp = gWs.LookUpMessage("{accepter='" & UserName & "' and is_affirmed<>'Y' order by send_time desc}")
            'If dsTmp Is Nothing Then
            '    Exit Sub
            'End If
            ''Me.LstBxMsg.DisplayMember = "message_content"
            ''Me.LstBxMsg.ValueMember = "serial_num"
            ''Me.LstBxMsg.DataSource = dsTmp.Tables(0)

            ''填充listview控件
            'Dim strTime As String

            'For i = 0 To dsTmp.Tables(0).Rows.Count - 1
            '    Me.lvwMessage.Items.Add(dsTmp.Tables(0).Rows(i).Item("serial_num"))
            '    Me.lvwMessage.Items(i).SubItems.Add(dsTmp.Tables(0).Rows(i).Item("send_time").ToString())
            '    If dsTmp.Tables(0).Rows(i).Item("project_code") Is DBNull.Value Then
            '        Me.lvwMessage.Items(i).SubItems.Add("")
            '    Else
            '        Me.lvwMessage.Items(i).SubItems.Add(dsTmp.Tables(0).Rows(i).Item("project_code"))
            '    End If

            '    Me.lvwMessage.Items(i).SubItems.Add(dsTmp.Tables(0).Rows(i).Item("message_content"))
            '    If dsTmp.Tables(0).Rows(i).Item("is_affirmed") = "N" Then
            '        lvwMessage.Items(i).Font = New Font(lvwMessage.Font, FontStyle.Bold)
            '    Else
            '        lvwMessage.Items(i).Font = New Font(lvwMessage.Font, FontStyle.Regular)
            '    End If
            'Next
        Catch ex As Exception
            'SWDialogBox.MessageBox.Show("*999", ex.Source, ex.Message, ex.StackTrace)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    'Private Sub lvwMessage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If Me.lvwMessage.Items.Count = 0 Then
    '        Exit Sub
    '    End If
    '    If Me.lvwMessage.FocusedItem Is Nothing Then
    '        Exit Sub
    '    End If

    '    Try
    '        Dim iNum As Int64
    '        Dim dsTmp As DataSet

    '        iNum = Me.lvwMessage.FocusedItem.Text

    '        dsTmp = gWs.LookUpMessage("{serial_num=" & iNum.ToString & "}")
    '        If Not dsTmp Is Nothing Then
    '            'MessageBox.Show(dsTmp.Tables(0).Rows(0).Item("message_content"), "消息内容", MessageBoxButtons.OK, MessageBoxIcon.None)
    '            dsTmp.Tables(0).Rows(0).Item("is_affirmed") = "R"
    '            gWs.UpdateMessage(dsTmp)
    '        End If
    '        If lvwMessage.FocusedItem.Font.Bold Then
    '            Me.lvwMessage.FocusedItem.Font = New Font(lvwMessage.Font, FontStyle.Regular)
    '        End If
    '    Catch ex As Exception
    '        SWDialogBox.MessageBox.Show("*999", ex.Source, ex.Message, ex.StackTrace)
    '    End Try
    'End Sub

    Private Sub FConsole_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
        If e.KeyCode = Keys.F5 Then
            Me.Refresh()
        End If
    End Sub

    Private Sub FMessage_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Refresh()

        tmrMessage.Interval = 300000
        tmrMessage.Enabled = True
    End Sub

    Private Sub tmrMessage_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrMessage.Tick
        Me.Refresh()
    End Sub

    Private Sub FConsole_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If MessageBox.Show("确认退出担保系统吗？ ", "退出系统", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Application.Exit()
        Else
            e.Cancel = True
        End If
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        Me.Refresh()
    End Sub
End Class


