

Public Class FStatisticsWorkingHours
    Inherits MainInterface.FQueryBase

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
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lblAttendPerson As System.Windows.Forms.Label
    Friend WithEvents dtpEnd As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpStart As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboStatisticsType As System.Windows.Forms.ComboBox
    Friend WithEvents cboStaff As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FStatisticsWorkingHours))
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lblAttendPerson = New System.Windows.Forms.Label()
        Me.dtpEnd = New System.Windows.Forms.DateTimePicker()
        Me.dtpStart = New System.Windows.Forms.DateTimePicker()
        Me.cboStaff = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboStatisticsType = New System.Windows.Forms.ComboBox()
        Me.grpCondition.SuspendLayout()
        CType(Me.grdTable, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'chkVisible
        '
        Me.chkVisible.BackColor = System.Drawing.Color.LightSkyBlue
        Me.chkVisible.Visible = True
        '
        'btnPrint
        '
        Me.btnPrint.BackColor = System.Drawing.Color.LightSkyBlue
        '
        'grpCondition
        '
        Me.grpCondition.BackColor = System.Drawing.Color.LightSkyBlue
        Me.grpCondition.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label1, Me.cboStatisticsType, Me.Label5, Me.Label4, Me.lblAttendPerson, Me.dtpEnd, Me.dtpStart, Me.cboStaff})
        Me.grpCondition.Visible = True
        '
        'btnExport
        '
        Me.btnExport.BackColor = System.Drawing.Color.LightSkyBlue
        Me.btnExport.Visible = True
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.Color.LightSkyBlue
        Me.btnExit.Visible = True
        '
        'grdTable
        '
        Me.grdTable.AccessibleName = "DataGrid"
        Me.grdTable.AccessibleRole = System.Windows.Forms.AccessibleRole.Table
        Me.grdTable.Visible = True
        '
        'btnRefresh
        '
        Me.btnRefresh.BackColor = System.Drawing.Color.LightSkyBlue
        Me.btnRefresh.Visible = True
        '
        'btnClear
        '
        Me.btnClear.BackColor = System.Drawing.Color.LightSkyBlue
        Me.btnClear.Visible = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Image = CType(resources.GetObject("Label5.Image"), System.Drawing.Bitmap)
        Me.Label5.Location = New System.Drawing.Point(456, 24)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(54, 14)
        Me.Label5.TabIndex = 23
        Me.Label5.Text = "截至日期"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Image = CType(resources.GetObject("Label4.Image"), System.Drawing.Bitmap)
        Me.Label4.Location = New System.Drawing.Point(248, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(54, 14)
        Me.Label4.TabIndex = 22
        Me.Label4.Text = "起始日期"
        '
        'lblAttendPerson
        '
        Me.lblAttendPerson.AutoSize = True
        Me.lblAttendPerson.Image = CType(resources.GetObject("lblAttendPerson.Image"), System.Drawing.Bitmap)
        Me.lblAttendPerson.Location = New System.Drawing.Point(8, 24)
        Me.lblAttendPerson.Name = "lblAttendPerson"
        Me.lblAttendPerson.Size = New System.Drawing.Size(54, 14)
        Me.lblAttendPerson.TabIndex = 20
        Me.lblAttendPerson.Text = "姓名"
        '
        'dtpEnd
        '
        Me.dtpEnd.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.dtpEnd.Location = New System.Drawing.Point(520, 21)
        Me.dtpEnd.Name = "dtpEnd"
        Me.dtpEnd.Size = New System.Drawing.Size(168, 21)
        Me.dtpEnd.TabIndex = 25
        '
        'dtpStart
        '
        Me.dtpStart.Location = New System.Drawing.Point(304, 21)
        Me.dtpStart.Name = "dtpStart"
        Me.dtpStart.Size = New System.Drawing.Size(144, 21)
        Me.dtpStart.TabIndex = 24
        '
        'cboStaff
        '
        Me.cboStaff.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboStaff.Location = New System.Drawing.Point(72, 20)
        Me.cboStaff.MaxLength = 4
        Me.cboStaff.Name = "cboStaff"
        Me.cboStaff.Size = New System.Drawing.Size(168, 22)
        Me.cboStaff.TabIndex = 21
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Image = CType(resources.GetObject("Label1.Image"), System.Drawing.Bitmap)
        Me.Label1.Location = New System.Drawing.Point(8, 64)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(54, 14)
        Me.Label1.TabIndex = 26
        Me.Label1.Text = "统计类型"
        '
        'cboStatisticsType
        '
        Me.cboStatisticsType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStatisticsType.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboStatisticsType.Items.AddRange(New Object() {"正常上班", "请假", "休假", "加班", "未记工作日志"})
        Me.cboStatisticsType.Location = New System.Drawing.Point(72, 56)
        Me.cboStatisticsType.MaxLength = 4
        Me.cboStatisticsType.Name = "cboStatisticsType"
        Me.cboStatisticsType.Size = New System.Drawing.Size(168, 22)
        Me.cboStatisticsType.TabIndex = 27
        '
        'FStatisticsWorkingHours
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.ClientSize = New System.Drawing.Size(712, 493)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnExport, Me.btnClear, Me.grdTable, Me.btnPrint, Me.btnExit, Me.btnRefresh, Me.chkVisible, Me.grpCondition})
        Me.Name = "FStatisticsWorkingHours"
        Me.Text = "工作工时统计"
        Me.grpCondition.ResumeLayout(False)
        CType(Me.grdTable, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Const ALLSTAFF As String = "<所有用户>"

    Private Function SetPermission(Optional ByVal userID As String = "") As Boolean
        If userID Is Nothing OrElse userID.Trim = String.Empty Then
            userID = UserName
        End If

        Dim i As Integer
        Dim dtStaff As DataTable = gWs.GetStaff("%").Tables(0)
        dtStaff.PrimaryKey = New DataColumn() {dtStaff.Columns("staff_name")}

        Dim findRow = dtStaff.Rows.Find(userID)
        If findRow Is Nothing Then
            MessageBox.Show("当前 " + userID + " 用户没有指定的权限定义。", "无效的登录用户", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return False
        End If

        Dim permission As String = findRow("read_logs_right").ToString()
        Dim department As String = findRow("dept_name").ToString()
        Dim departmentList As String = SplitStringListForWhereCondition(findRow("logs_department_list").ToString())

        cboStaff.Items.Clear()
        cboStaff.Items.Add(ALLSTAFF)

        Select Case permission
            Case "0"
                cboStaff.Enabled = True
                cboStaff.DropDownStyle = ComboBoxStyle.DropDown

                For i = 0 To dtStaff.Rows.Count - 1
                    cboStaff.Items.Add(dtStaff.Rows(i)("staff_name"))
                Next
            Case "2"
                cboStaff.Enabled = True
                cboStaff.DropDownStyle = ComboBoxStyle.DropDownList

                Dim rows() As DataRow = dtStaff.Select("dept_name = '" + department + "'")
                For i = 0 To rows.Length - 1
                    cboStaff.Items.Add(rows(i)("staff_name"))
                Next
            Case "3"
                cboStaff.Enabled = True
                cboStaff.DropDownStyle = ComboBoxStyle.DropDownList

                cboStaff.Items.Add(userID)

                Dim rows() As DataRow = dtStaff.Select("read_logs_right = '1'")
                For i = 0 To rows.Length - 1
                    cboStaff.Items.Add(rows(i)("staff_name"))
                Next
            Case "4"
                cboStaff.Enabled = True
                cboStaff.DropDownStyle = ComboBoxStyle.DropDownList


                Dim rows() As DataRow = dtStaff.Select("dept_name in (" + departmentList + ")")
                For i = 0 To rows.Length - 1
                    cboStaff.Items.Add(rows(i)("staff_name"))
                Next
            Case Else
                cboStaff.Enabled = False
                cboStaff.DropDownStyle = ComboBoxStyle.Simple
        End Select

        cboStaff.Text = userID

        Return True
    End Function

    Protected Overrides Sub Clear()
        If Not Me.IsHandleCreated Then
            Return
        End If

        cboStatisticsType.SelectedIndex = 0
    End Sub

    Protected Overloads Sub GetCondition(ByRef queryType As String, ByRef dateStart As DateTime, ByRef dateEnd As DateTime, ByRef attendPerson As String, ByRef period As String)
        queryType = cboStatisticsType.SelectedIndex.ToString
        dateStart = dtpStart.Value.Date
        dateEnd = dtpEnd.Value.AddDays(1).Date

        If cboStaff.Text.Trim() <> "" And cboStaff.Text <> ALLSTAFF Then
            attendPerson = cboStaff.Text.Trim
        Else
            Dim i As Integer
            Dim sql As String

            For i = 1 To cboStaff.Items.Count - 1
                sql += IIf(sql = "", cboStaff.Items(i).ToString, "," & cboStaff.Items(i).ToString)
            Next
            attendPerson = sql.Trim
        End If

        period = String.Empty
    End Sub

    Private Function ConditionIsEffective() As Boolean
        If dtpStart.Value.Date > dtpEnd.Value.Date Then
            SWDialogBox.MessageBox.Show("$1008", "开始日期", "截止日期")
            dtpEnd.Focus()
            Return False
        End If
        If cboStaff.Text = "" Then
            SWDialogBox.MessageBox.Show("$1001", "员工姓名")
            Return False
        End If
        If cboStatisticsType.SelectedItem Is Nothing Then
            SWDialogBox.MessageBox.Show("$1001", "统计类型")
            Return False
        End If
        Return True
    End Function

    Protected Overloads Overrides Sub Refresh(ByVal condition As String)
        If Not ConditionIsEffective() Then
            Return
        End If

        Dim dateStart, dateEnd As System.DateTime
        Dim attendPerson, queryType, period As String

        GetCondition(queryType, dateStart, dateEnd, attendPerson, period)

        Dim dsResult As DataSet = gWs.GetWorkingHours(attendPerson, dateStart, dateEnd, period, queryType)

        dsResult.Tables("Table").TableName = "TQueryWorkingHours"
        grdTable.SetDataBinding(dsResult, "TQueryWorkingHours")
        AddTableStyle(queryType)

        Me.Text = "工时统计(" + dsResult.Tables("TQueryWorkingHours").Rows.Count.ToString + ")"
    End Sub

    Private Function CreateOneColumn(ByVal mappingName As String, ByVal headertext As String, ByVal width As Int32, Optional ByVal format As String = "")
        Dim col As DataGridTextBoxColumn = New DataGridTextBoxColumn()
        col.MappingName = mappingName
        col.HeaderText = headertext
        col.Width = width
        col.NullText = String.Empty
        col.Format = format
        Return col
    End Function

    Private Sub AddTableStyle(ByVal QueryType As Int16)
        Dim dgts As DataGridTableStyle = New DataGridTableStyle()
        dgts.MappingName = "TQueryWorkingHours"
        Select Case QueryType
            Case 0
                dgts.GridColumnStyles.AddRange(New DataGridColumnStyle() {CreateOneColumn("StaffName", "员工姓名", 80), _
                                                                          CreateOneColumn("RealWorkHours", "正常上班工时", 120), _
                                                                          CreateOneColumn("AbsentHours", "未记日志工时", 120), _
                                                                          CreateOneColumn("saturation", "饱和度(%)", 100) _
                                                                          })
            Case 1
                dgts.GridColumnStyles.AddRange(New DataGridColumnStyle() {CreateOneColumn("StaffName", "员工姓名", 80), _
                                                                          CreateOneColumn("LeaveHours", "请假工时", 120) _
                                                                          })
            Case 2
                dgts.GridColumnStyles.AddRange(New DataGridColumnStyle() {CreateOneColumn("StaffName", "员工姓名", 80), _
                                                                          CreateOneColumn("VacationHours", "休假工时", 120) _
                                                                          })
            Case 3
                dgts.GridColumnStyles.AddRange(New DataGridColumnStyle() {CreateOneColumn("StaffName", "员工姓名", 80), _
                                                                          CreateOneColumn("OverHours", "加班工时", 120), _
                                                                          CreateOneColumn("WorkingHours", "正常上班工时", 120), _
                                                                          CreateOneColumn("AbsentHours", "未记日志工时", 120), _
                                                                          CreateOneColumn("saturation", "饱和度(%)", 100) _
                                                                          })
            Case 4
                dgts.GridColumnStyles.AddRange(New DataGridColumnStyle() {CreateOneColumn("StaffName", "员工姓名", 80), _
                                                                          CreateOneColumn("AbsentHours", "未记日志工时", 120) _
                                                                          })
        End Select
        Me.grdTable.TableStyles.Clear()
        Me.grdTable.TableStyles.Add(dgts)
    End Sub

    Private Sub FStatisticsWorkingHours_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim SysDate As DateTime = gWs.GetSysTime
        dtpEnd.Value = SysDate.Date
        dtpStart.Value = SysDate.Date

        If Not SetPermission() Then
            Me.btnPrint.Enabled = False
            Me.btnRefresh.Enabled = False
        End If
        Me.cboStatisticsType.SelectedIndex = 0
    End Sub

    Protected Overrides Sub Export()
        ReportToExcel.ToExcel(grdTable, cboStatisticsType.SelectedItem.ToString, cboStatisticsType.SelectedItem.ToString)
    End Sub
End Class
