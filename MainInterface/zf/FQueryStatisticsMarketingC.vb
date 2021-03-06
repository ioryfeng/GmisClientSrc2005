Imports System
Imports System.Data
Imports System.Data.SqlClient

Public Class FQueryStatisticsMarketingC
    Inherits FQueryBase

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
	Friend WithEvents lblDateStart As System.Windows.Forms.Label
	Friend WithEvents cboYearStart As System.Windows.Forms.ComboBox
	Friend WithEvents cboYearEnd As System.Windows.Forms.ComboBox
	Friend WithEvents lblDateEnd As System.Windows.Forms.Label
	Friend WithEvents cboMonthEnd As System.Windows.Forms.ComboBox
	Friend WithEvents tsTable As System.Windows.Forms.DataGridTableStyle
	Friend WithEvents csYear As System.Windows.Forms.DataGridTextBoxColumn
	Friend WithEvents csMonth As System.Windows.Forms.DataGridTextBoxColumn
	Friend WithEvents csMaxSum As System.Windows.Forms.DataGridTextBoxColumn
	Friend WithEvents csMinSum As System.Windows.Forms.DataGridTextBoxColumn
	Friend WithEvents csAvgSum As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents csYearMonth As System.Windows.Forms.DataGridTextBoxColumn
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblDateStart = New System.Windows.Forms.Label()
        Me.cboYearStart = New System.Windows.Forms.ComboBox()
        Me.cboYearEnd = New System.Windows.Forms.ComboBox()
        Me.lblDateEnd = New System.Windows.Forms.Label()
        Me.cboMonthEnd = New System.Windows.Forms.ComboBox()
        Me.tsTable = New System.Windows.Forms.DataGridTableStyle()
        Me.csYearMonth = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.csYear = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.csMonth = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.csMaxSum = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.csMinSum = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.csAvgSum = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.grpCondition.SuspendLayout()
        CType(Me.grdTable, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'chkVisible
        '
        Me.chkVisible.Location = New System.Drawing.Point(8, 80)
        Me.chkVisible.Visible = True
        '
        'grpCondition
        '
        Me.grpCondition.Controls.AddRange(New System.Windows.Forms.Control() {Me.cboMonthEnd, Me.cboYearEnd, Me.lblDateEnd, Me.cboYearStart, Me.lblDateStart})
        Me.grpCondition.Size = New System.Drawing.Size(696, 56)
        Me.grpCondition.Visible = True
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(616, 80)
        Me.btnExit.Visible = True
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(192, 80)
        Me.btnClear.Visible = True
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(520, 80)
        Me.btnPrint.Visible = True
        '
        'btnRefresh
        '
        Me.btnRefresh.Location = New System.Drawing.Point(424, 80)
        Me.btnRefresh.Visible = True
        '
        'btnExport
        '
        Me.btnExport.Location = New System.Drawing.Point(520, 80)
        Me.btnExport.Visible = True
        '
        'grdTable
        '
        Me.grdTable.AccessibleName = "DataGrid"
        Me.grdTable.AccessibleRole = System.Windows.Forms.AccessibleRole.Table
        Me.grdTable.Location = New System.Drawing.Point(8, 112)
        Me.grdTable.Size = New System.Drawing.Size(696, 376)
        Me.grdTable.TableStyles.AddRange(New System.Windows.Forms.DataGridTableStyle() {Me.tsTable})
        Me.grdTable.Visible = True
        '
        'lblDateStart
        '
        Me.lblDateStart.Location = New System.Drawing.Point(16, 32)
        Me.lblDateStart.Name = "lblDateStart"
        Me.lblDateStart.Size = New System.Drawing.Size(64, 16)
        Me.lblDateStart.TabIndex = 0
        Me.lblDateStart.Text = "起始年份"
        '
        'cboYearStart
        '
        Me.cboYearStart.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboYearStart.Location = New System.Drawing.Point(80, 24)
        Me.cboYearStart.MaxLength = 4
        Me.cboYearStart.Name = "cboYearStart"
        Me.cboYearStart.Size = New System.Drawing.Size(56, 22)
        Me.cboYearStart.TabIndex = 1
        '
        'cboYearEnd
        '
        Me.cboYearEnd.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboYearEnd.Location = New System.Drawing.Point(248, 24)
        Me.cboYearEnd.MaxLength = 4
        Me.cboYearEnd.Name = "cboYearEnd"
        Me.cboYearEnd.Size = New System.Drawing.Size(56, 22)
        Me.cboYearEnd.TabIndex = 3
        '
        'lblDateEnd
        '
        Me.lblDateEnd.Location = New System.Drawing.Point(184, 32)
        Me.lblDateEnd.Name = "lblDateEnd"
        Me.lblDateEnd.Size = New System.Drawing.Size(64, 16)
        Me.lblDateEnd.TabIndex = 2
        Me.lblDateEnd.Text = "截止年月"
        '
        'cboMonthEnd
        '
        Me.cboMonthEnd.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboMonthEnd.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboMonthEnd.Location = New System.Drawing.Point(312, 24)
        Me.cboMonthEnd.MaxLength = 2
        Me.cboMonthEnd.Name = "cboMonthEnd"
        Me.cboMonthEnd.Size = New System.Drawing.Size(40, 22)
        Me.cboMonthEnd.TabIndex = 4
        '
        'tsTable
        '
        Me.tsTable.DataGrid = Me.grdTable
        Me.tsTable.GridColumnStyles.AddRange(New System.Windows.Forms.DataGridColumnStyle() {Me.csYearMonth, Me.csYear, Me.csMonth, Me.csMaxSum, Me.csMinSum, Me.csAvgSum})
        Me.tsTable.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.tsTable.MappingName = "TQueryStatisticsMarketingC"
        '
        'csYearMonth
        '
        Me.csYearMonth.Format = ""
        Me.csYearMonth.FormatInfo = Nothing
        Me.csYearMonth.HeaderText = "年月"
        Me.csYearMonth.MappingName = "YearMonth"
        Me.csYearMonth.Width = 80
        '
        'csYear
        '
        Me.csYear.Format = ""
        Me.csYear.FormatInfo = Nothing
        Me.csYear.HeaderText = "年"
        Me.csYear.MappingName = "yearPart"
        Me.csYear.Width = 0
        '
        'csMonth
        '
        Me.csMonth.Format = ""
        Me.csMonth.FormatInfo = Nothing
        Me.csMonth.HeaderText = "月"
        Me.csMonth.MappingName = "monthPart"
        Me.csMonth.Width = 0
        '
        'csMaxSum
        '
        Me.csMaxSum.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.csMaxSum.Format = "N"
        Me.csMaxSum.FormatInfo = Nothing
        Me.csMaxSum.HeaderText = "单笔最大金额"
        Me.csMaxSum.MappingName = "maxSum"
        Me.csMaxSum.Width = 75
        '
        'csMinSum
        '
        Me.csMinSum.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.csMinSum.Format = "N"
        Me.csMinSum.FormatInfo = Nothing
        Me.csMinSum.HeaderText = "单笔最小金额"
        Me.csMinSum.MappingName = "minSum"
        Me.csMinSum.Width = 75
        '
        'csAvgSum
        '
        Me.csAvgSum.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.csAvgSum.Format = "N"
        Me.csAvgSum.FormatInfo = Nothing
        Me.csAvgSum.HeaderText = "平均每笔金额"
        Me.csAvgSum.MappingName = "avgSum"
        Me.csAvgSum.Width = 75
        '
        'FQueryStatisticsMarketingC
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.ClientSize = New System.Drawing.Size(712, 493)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnExport, Me.btnClear, Me.grdTable, Me.btnPrint, Me.btnExit, Me.btnRefresh, Me.chkVisible, Me.grpCondition})
        Me.Name = "FQueryStatisticsMarketingC"
        Me.Text = "担保金额有关统计"
        Me.grpCondition.ResumeLayout(False)
        CType(Me.grdTable, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

	Protected Overloads Sub GetCondition(ByRef dateStart As DateTime, ByRef dateEnd As DateTime)
		'yearStart = MyBase.ToNumeric(cboYearStart.Text)
		'yearEnd = MyBase.ToNumeric(cboYearEnd.Text)
		'monthEnd = MyBase.ToNumeric(cboMonthEnd.Text)
		'cboYearStart.Text = yearStart.ToString()
		'cboYearEnd.Text = yearEnd.ToString()
		'cboMonthEnd.Text = monthEnd.ToString()

		Dim yearStart, yearEnd, monthEnd As Int32

		Try
			yearStart = Int32.Parse(cboYearStart.Text)
		Catch
			cboYearStart.Text = System.DateTime.Today.Year.ToString()
			yearStart = System.DateTime.Today.Year
		End Try
		dateStart = New System.DateTime(yearstart, 1, 1)

		Try
			yearEnd = Int32.Parse(cboYearEnd.Text)
		Catch
			cboYearEnd.Text = System.DateTime.Today.Year.ToString()
			yearEnd = System.DateTime.Today.Year
		End Try

		Try
			monthEnd = Int32.Parse(cboMonthEnd.Text)
		Catch
			cboMonthEnd.Text = System.DateTime.Today.Month.ToString()
			monthEnd = System.DateTime.Today.Month
		End Try

		dateEnd = New System.DateTime(yearEnd, monthEnd, 1)
	End Sub

	Protected Overloads Overrides Sub Refresh(ByVal condition As String)
		Dim dateStart, dateEnd As System.DateTime

		GetCondition(dateStart, dateEnd)

		Dim dsResult As DataSet = gWs.PQueryStatisticsMarketingC(dateStart, dateEnd)

		dsResult.Tables("Table").TableName = "TQueryStatisticsMarketingC"

        Dim maxSum, minSum, avgSum As Double
        Dim row As DataRow
        For Each row In dsResult.Tables("TQueryStatisticsMarketingC").Rows
            If Not row.IsNull("maxSum") AndAlso row("maxSum") > maxSum Then
                maxSum = row("maxSum")
            End If
            If Not row.IsNull("minSum") AndAlso row("minSum") < minSum Then
                minSum = row("minSum")
            End If
            If Not row.IsNull("avgSum") Then
                avgSum += row("avgSum")
            End If
        Next

        Dim newRow As DataRow = dsResult.Tables("TQueryStatisticsMarketingC").NewRow()
        newRow("yearPart") = "合计："
        newRow("maxSum") = maxSum
        newRow("minSum") = minSum
        If dsResult.Tables("TQueryStatisticsMarketingC").Rows.Count > 0 Then
            newRow("avgSum") = avgSum / dsResult.Tables("TQueryStatisticsMarketingC").Rows.Count
        End If
        dsResult.Tables("TQueryStatisticsMarketingC").Rows.Add(newRow)

        dsResult.Tables("TQueryStatisticsMarketingC").Columns.Add("YearMonth", GetType(String), "yearPart + IIF(LEN(ISNULL(monthPart, '')) = 1, '0', '') + ISNULL(monthPart, '')")

        MyBase.DataSource = dsResult
        grdTable.DataSource = dsResult
        grdTable.SetDataBinding(dsResult, "TQueryStatisticsMarketingC")
        Me.Text = "担保金额有关统计(" + dsResult.Tables("TQueryStatisticsMarketingC").Rows.Count.ToString + ")"
    End Sub

    Private Sub FQueryStatisticsMarketingC_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim i As Int32 = 0

        For i = 1990 To 2050
            cboYearEnd.Items.Add(i.ToString())
            cboYearStart.Items.Add(i.ToString())
        Next
        For i = 1 To 12
            cboMonthEnd.Items.Add(i.ToString("00"))
        Next

        cboYearEnd.SelectedIndex = System.DateTime.Today.Year - 1990
        cboYearStart.SelectedIndex = System.DateTime.Today.Year - 1990
        cboMonthEnd.SelectedIndex = System.DateTime.Today.Month - 1

        MyBase.ReportTitle = "担保金额有关统计"
        MyBase.ReportFile = Application.StartupPath + "\Reports\QueryStatisticsMarketingC.rpt"
    End Sub

    Protected Overrides Sub Clear()
        If Not Me.IsHandleCreated Then
            Return
        End If

        cboYearEnd.Text = System.DateTime.Today.Year.ToString()
        cboYearStart.Text = System.DateTime.Today.Year.ToString()
        If cboMonthEnd.Items.Count > 0 Then
            cboMonthEnd.SelectedIndex = System.DateTime.Today.Month - 1
        End If
    End Sub

    Private Sub ComboBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cboYearStart.KeyPress, cboMonthEnd.KeyPress, cboYearEnd.KeyPress
        e.Handled = True

        If (Not Char.IsDigit(e.KeyChar)) And Asc(e.KeyChar) <> Keys.Back Then
            e.Handled = True
        End If
    End Sub
End Class
