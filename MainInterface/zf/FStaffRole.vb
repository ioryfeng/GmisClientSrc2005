Imports System
Imports System.Data
Imports System.Data.SqlClient

Public Class FStaffRole
    Inherits FBaseData

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
	Friend WithEvents tsStaffRole As System.Windows.Forms.DataGridTableStyle
	Friend WithEvents csRoleID As DataGridComboBoxColumn
    Friend WithEvents csStaffName As DataGridComboBoxColumn
    Friend WithEvents txtConferenceRoom As DataGridTextBoxColumn
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.tsStaffRole = New System.Windows.Forms.DataGridTableStyle()
		Me.csRoleID = New DataGridComboBoxColumn()
        Me.csStaffName = New DataGridComboBoxColumn
        Me.txtConferenceRoom = New DataGridTextBoxColumn
		Me.SuspendLayout()
		'
		'tsStaffRole
		'
		Me.tsStaffRole.DataGrid = Me.grdTable
        Me.tsStaffRole.GridColumnStyles.AddRange(New System.Windows.Forms.DataGridColumnStyle() {Me.csRoleID, Me.csStaffName, Me.txtConferenceRoom})
		Me.tsStaffRole.HeaderForeColor = System.Drawing.SystemColors.ControlText
		Me.tsStaffRole.MappingName = "TStaffRole"
		'
		'csRoleID
		'
		Me.csRoleID.Format = ""
		Me.csRoleID.FormatInfo = Nothing
		Me.csRoleID.HeaderText = "所属角色"
		Me.csRoleID.MappingName = "role_id"
		Me.csRoleID.Width = 120
		'
		'csStaffName
		'
		Me.csStaffName.Format = ""
		Me.csStaffName.FormatInfo = Nothing
		Me.csStaffName.HeaderText = "员工姓名"
		Me.csStaffName.MappingName = "staff_name"
        Me.csStaffName.Width = 120


        'txtConferenceRoom
        '
        Me.txtConferenceRoom.Format = ""
        Me.txtConferenceRoom.FormatInfo = Nothing
        Me.txtConferenceRoom.HeaderText = "会场"
        Me.txtConferenceRoom.MappingName = "conference_address"
        Me.txtConferenceRoom.Width = 200

		'
		'FStaffRole
		'
		Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
		Me.ClientSize = New System.Drawing.Size(512, 333)
		Me.Name = "FStaffRole"
		Me.Text = "员工角色设置"
		Me.ResumeLayout(False)

	End Sub

#End Region

	Public Overloads Overrides Sub Refresh()
		Try
			Me.Cursor = Cursors.WaitCursor

			Dim dsResult As DataSet

			dsResult = gWs.GetStaffRole("%", "%")
			grdTable.SetDataBinding(dsResult, "TStaffRole")
		Catch ex As System.Exception
			'统一错误处理
			ExceptionHandler.ShowMessageBox(ex)
		Finally
			Me.Cursor = Cursors.Default
		End Try
	End Sub

	Public Overrides Function Update() As Boolean
		Try
			Me.Cursor = Cursors.WaitCursor

            Dim dsCommit As DataSet = CType(grdTable.DataSource, DataSet)

            If dsCommit Is Nothing OrElse (Not dsCommit.HasChanges()) Then
                Return False
            End If

            Dim result As String = gWs.UpdateStaffRole(dsCommit.GetChanges())

            If result = "1" Then
                '接受所有更改
                dsCommit.AcceptChanges()

                Return True
            Else
                '显示服务器的返回错误信息
                SWDialogBox.MessageBox.Show("*999", Me.Name + ".Update", result, "")
            End If
        Catch ex As System.Exception
			'统一错误处理
			ExceptionHandler.ShowMessageBox(ex)
		Finally
			Me.Cursor = Cursors.Default
		End Try

		Return False
	End Function

	Private Sub FStaffRole_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
		Dim dsRole As DataSet = gWs.GetRole("%")
        Dim dsStaff As DataSet = gWs.GetStaff("%")
        Dim dsConferenceRoom As DataSet = gWs.FetchConfernceRoom("%")

		csRoleID.ColumnComboBox.DataSource = dsRole.Tables("TRole").DefaultView
		csRoleID.ColumnComboBox.DisplayMember = "role_name"
		csRoleID.ColumnComboBox.ValueMember = "role_id"

		csStaffName.ColumnComboBox.DataSource = dsStaff.Tables("TStaff").DefaultView
		csStaffName.ColumnComboBox.DisplayMember = "staff_name"
        csStaffName.ColumnComboBox.ValueMember = "staff_name"

        'tsStaffRole.PreferredRowHeight = csRoleID.ColumnComboBox.Height + 2
		grdTable.TableStyles.Add(Me.tsStaffRole)
	End Sub
End Class
