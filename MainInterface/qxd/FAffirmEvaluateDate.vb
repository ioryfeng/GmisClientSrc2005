

Public Class FAffirmEvaluateDate
    Inherits MainInterface.frmOppGuarantee

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
    Friend WithEvents corp_contact_phone As System.Windows.Forms.TextBox
    Friend WithEvents corp_contact_person As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents dtpEvaluateAffirmDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnViewPic As System.Windows.Forms.Button
    Friend WithEvents btnCommit As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents DataGridTableStyle2 As System.Windows.Forms.DataGridTableStyle
    Friend WithEvents DataGridTextBoxColumn6 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn7 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn8 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn9 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn3 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn10 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn11 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn14 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtManagerA As System.Windows.Forms.TextBox
    Friend WithEvents dtpBooking As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents DataGridTextBoxColumn12 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn13 As System.Windows.Forms.DataGridTextBoxColumn
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FAffirmEvaluateDate))
        Me.corp_contact_phone = New System.Windows.Forms.TextBox()
        Me.corp_contact_person = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.dtpEvaluateAffirmDate = New System.Windows.Forms.DateTimePicker()
        Me.btnViewPic = New System.Windows.Forms.Button()
        Me.btnCommit = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.DataGridTableStyle2 = New System.Windows.Forms.DataGridTableStyle()
        Me.DataGridTextBoxColumn6 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn7 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn13 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn12 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn8 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn3 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn10 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn11 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn9 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn14 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.txtManagerA = New System.Windows.Forms.TextBox()
        Me.dtpBooking = New System.Windows.Forms.DateTimePicker()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.dgOppGuarantee, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label6
        '
        Me.Label6.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(216, 28)
        Me.Label6.Size = New System.Drawing.Size(66, 14)
        Me.Label6.Visible = True
        '
        'Label7
        '
        Me.Label7.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.Label7.Visible = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Location = New System.Drawing.Point(8, 376)
        Me.GroupBox4.Size = New System.Drawing.Size(680, 112)
        Me.GroupBox4.Visible = True
        '
        'Label11
        '
        Me.Label11.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.Label11.Visible = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.GroupBox3.Location = New System.Drawing.Point(392, 40)
        Me.GroupBox3.Size = New System.Drawing.Size(296, 264)
        Me.GroupBox3.Visible = True
        '
        'Label10
        '
        Me.Label10.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.Label10.Visible = True
        '
        'txtRemark
        '
        Me.txtRemark.Anchor = ((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.txtRemark.ReadOnly = True
        Me.txtRemark.Size = New System.Drawing.Size(600, 21)
        Me.txtRemark.Visible = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.GroupBox2.Size = New System.Drawing.Size(376, 264)
        Me.GroupBox2.Visible = True
        '
        'dgOppGuarantee
        '
        Me.dgOppGuarantee.AccessibleName = "DataGrid"
        Me.dgOppGuarantee.AccessibleRole = System.Windows.Forms.AccessibleRole.Table
        Me.dgOppGuarantee.Size = New System.Drawing.Size(370, 244)
        Me.dgOppGuarantee.TableStyles.AddRange(New System.Windows.Forms.DataGridTableStyle() {Me.DataGridTableStyle2})
        Me.dgOppGuarantee.Visible = True
        '
        'dateEvaluate
        '
        Me.dateEvaluate.Anchor = ((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.dateEvaluate.Enabled = False
        Me.dateEvaluate.Size = New System.Drawing.Size(176, 21)
        Me.dateEvaluate.Value = New Date(2004, 6, 10, 9, 21, 12, 921)
        Me.dateEvaluate.Visible = True
        '
        'cmbType
        '
        Me.cmbType.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.cmbType.Enabled = False
        Me.cmbType.ItemHeight = 12
        Me.cmbType.Visible = True
        '
        'txtContract
        '
        Me.txtContract.Anchor = ((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.txtContract.ReadOnly = True
        Me.txtContract.Size = New System.Drawing.Size(176, 21)
        Me.txtContract.Visible = True
        '
        'txtGuarantee
        '
        Me.txtGuarantee.Anchor = ((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.txtGuarantee.ReadOnly = True
        Me.txtGuarantee.Size = New System.Drawing.Size(176, 21)
        Me.txtGuarantee.Visible = True
        '
        'cmbMeasure
        '
        Me.cmbMeasure.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.cmbMeasure.DropDownWidth = 120
        Me.cmbMeasure.Enabled = False
        Me.cmbMeasure.ItemHeight = 12
        Me.cmbMeasure.Location = New System.Drawing.Point(288, 25)
        Me.cmbMeasure.Size = New System.Drawing.Size(120, 20)
        Me.cmbMeasure.Visible = True
        '
        'txtEvaluate
        '
        Me.txtEvaluate.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.txtEvaluate.Location = New System.Drawing.Point(288, 81)
        Me.txtEvaluate.ReadOnly = True
        Me.txtEvaluate.Size = New System.Drawing.Size(120, 21)
        Me.txtEvaluate.Visible = True
        '
        'Label3
        '
        Me.Label3.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.Label3.Visible = True
        '
        'Label4
        '
        Me.Label4.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.Label4.Visible = True
        '
        'dgDetail
        '
        Me.dgDetail.AccessibleName = "DataGrid"
        Me.dgDetail.AccessibleRole = System.Windows.Forms.AccessibleRole.Table
        Me.dgDetail.Size = New System.Drawing.Size(290, 244)
        Me.dgDetail.Visible = True
        '
        'txtOriginal
        '
        Me.txtOriginal.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.txtOriginal.ReadOnly = True
        Me.txtOriginal.Visible = True
        '
        'Label5
        '
        Me.Label5.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.Label5.Visible = True
        '
        'Label9
        '
        Me.Label9.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.Label9.Location = New System.Drawing.Point(440, 56)
        Me.Label9.Visible = True
        '
        'txtName
        '
        Me.txtName.Size = New System.Drawing.Size(368, 21)
        Me.txtName.Visible = True
        '
        'ImageListBasic
        '
        Me.ImageListBasic.ImageStream = CType(resources.GetObject("ImageListBasic.ImageStream"), System.Windows.Forms.ImageListStreamer)
        '
        'btnExit
        '
        Me.btnExit.Image = CType(resources.GetObject("btnExit.Image"), System.Drawing.Bitmap)
        Me.btnExit.ImageIndex = -1
        Me.btnExit.ImageList = Nothing
        Me.btnExit.Location = New System.Drawing.Point(451, 500)
        Me.btnExit.Size = New System.Drawing.Size(77, 24)
        Me.btnExit.Visible = True
        '
        'corp_contact_phone
        '
        Me.corp_contact_phone.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.corp_contact_phone.Location = New System.Drawing.Point(312, 320)
        Me.corp_contact_phone.Name = "corp_contact_phone"
        Me.corp_contact_phone.ReadOnly = True
        Me.corp_contact_phone.Size = New System.Drawing.Size(152, 21)
        Me.corp_contact_phone.TabIndex = 26
        Me.corp_contact_phone.Text = ""
        '
        'corp_contact_person
        '
        Me.corp_contact_person.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.corp_contact_person.Location = New System.Drawing.Point(88, 320)
        Me.corp_contact_person.Name = "corp_contact_person"
        Me.corp_contact_person.ReadOnly = True
        Me.corp_contact_person.Size = New System.Drawing.Size(128, 21)
        Me.corp_contact_person.TabIndex = 25
        Me.corp_contact_person.Text = ""
        '
        'Label12
        '
        Me.Label12.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(232, 323)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(54, 14)
        Me.Label12.TabIndex = 24
        Me.Label12.Text = "联系电话"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label13
        '
        Me.Label13.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(8, 323)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(66, 14)
        Me.Label13.TabIndex = 23
        Me.Label13.Text = "企业联系人"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label8
        '
        Me.Label8.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(232, 352)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(79, 14)
        Me.Label8.TabIndex = 27
        Me.Label8.Text = "确认评估时间"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'dtpEvaluateAffirmDate
        '
        Me.dtpEvaluateAffirmDate.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.dtpEvaluateAffirmDate.Location = New System.Drawing.Point(312, 349)
        Me.dtpEvaluateAffirmDate.Name = "dtpEvaluateAffirmDate"
        Me.dtpEvaluateAffirmDate.Size = New System.Drawing.Size(152, 21)
        Me.dtpEvaluateAffirmDate.TabIndex = 28
        '
        'btnViewPic
        '
        Me.btnViewPic.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnViewPic.Enabled = False
        Me.btnViewPic.Image = CType(resources.GetObject("btnViewPic.Image"), System.Drawing.Bitmap)
        Me.btnViewPic.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnViewPic.ImageIndex = 17
        Me.btnViewPic.ImageList = Me.ImageListBasic
        Me.btnViewPic.Location = New System.Drawing.Point(171, 500)
        Me.btnViewPic.Name = "btnViewPic"
        Me.btnViewPic.Size = New System.Drawing.Size(93, 24)
        Me.btnViewPic.TabIndex = 29
        Me.btnViewPic.Text = "浏览图片(&V)"
        Me.btnViewPic.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnCommit
        '
        Me.btnCommit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCommit.Image = CType(resources.GetObject("btnCommit.Image"), System.Drawing.Bitmap)
        Me.btnCommit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCommit.ImageIndex = 11
        Me.btnCommit.ImageList = Me.ImageListBasic
        Me.btnCommit.Location = New System.Drawing.Point(363, 500)
        Me.btnCommit.Name = "btnCommit"
        Me.btnCommit.Size = New System.Drawing.Size(77, 24)
        Me.btnCommit.TabIndex = 30
        Me.btnCommit.Text = "提 交(&C)"
        Me.btnCommit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnSave
        '
        Me.btnSave.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnSave.Image = CType(resources.GetObject("btnSave.Image"), System.Drawing.Bitmap)
        Me.btnSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnSave.ImageIndex = 10
        Me.btnSave.ImageList = Me.ImageListBasic
        Me.btnSave.Location = New System.Drawing.Point(275, 500)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(77, 24)
        Me.btnSave.TabIndex = 31
        Me.btnSave.Text = "保 存(&S)"
        Me.btnSave.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'DataGridTableStyle2
        '
        Me.DataGridTableStyle2.AllowSorting = False
        Me.DataGridTableStyle2.DataGrid = Me.dgOppGuarantee
        Me.DataGridTableStyle2.GridColumnStyles.AddRange(New System.Windows.Forms.DataGridColumnStyle() {Me.DataGridTextBoxColumn6, Me.DataGridTextBoxColumn7, Me.DataGridTextBoxColumn13, Me.DataGridTextBoxColumn12, Me.DataGridTextBoxColumn8, Me.DataGridTextBoxColumn3, Me.DataGridTextBoxColumn10, Me.DataGridTextBoxColumn11, Me.DataGridTextBoxColumn9, Me.DataGridTextBoxColumn14})
        Me.DataGridTableStyle2.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridTableStyle2.MappingName = "opposite_guarantee"
        '
        'DataGridTextBoxColumn6
        '
        Me.DataGridTextBoxColumn6.Format = ""
        Me.DataGridTextBoxColumn6.FormatInfo = Nothing
        Me.DataGridTextBoxColumn6.HeaderText = "类型"
        Me.DataGridTextBoxColumn6.MappingName = "type_name"
        Me.DataGridTextBoxColumn6.NullText = ""
        Me.DataGridTextBoxColumn6.ReadOnly = True
        Me.DataGridTextBoxColumn6.Width = 75
        '
        'DataGridTextBoxColumn7
        '
        Me.DataGridTextBoxColumn7.Format = "yyyy年MM月dd日"
        Me.DataGridTextBoxColumn7.FormatInfo = Nothing
        Me.DataGridTextBoxColumn7.HeaderText = "预约评估日期"
        Me.DataGridTextBoxColumn7.MappingName = "evaluate_date"
        Me.DataGridTextBoxColumn7.NullText = ""
        Me.DataGridTextBoxColumn7.ReadOnly = True
        Me.DataGridTextBoxColumn7.Width = 95
        '
        'DataGridTextBoxColumn13
        '
        Me.DataGridTextBoxColumn13.Format = "yyyy年MM月dd日"
        Me.DataGridTextBoxColumn13.FormatInfo = Nothing
        Me.DataGridTextBoxColumn13.HeaderText = "确认评估日期"
        Me.DataGridTextBoxColumn13.MappingName = "affirm_evaluate_date"
        Me.DataGridTextBoxColumn13.NullText = ""
        Me.DataGridTextBoxColumn13.Width = 95
        '
        'DataGridTextBoxColumn12
        '
        Me.DataGridTextBoxColumn12.Format = ""
        Me.DataGridTextBoxColumn12.FormatInfo = Nothing
        Me.DataGridTextBoxColumn12.HeaderText = "评估师"
        Me.DataGridTextBoxColumn12.MappingName = "Evaluate_person"
        Me.DataGridTextBoxColumn12.NullText = ""
        Me.DataGridTextBoxColumn12.Width = 75
        '
        'DataGridTextBoxColumn8
        '
        Me.DataGridTextBoxColumn8.Format = ""
        Me.DataGridTextBoxColumn8.FormatInfo = Nothing
        Me.DataGridTextBoxColumn8.HeaderText = "反担保方式"
        Me.DataGridTextBoxColumn8.MappingName = "opposite_guarantee_form"
        Me.DataGridTextBoxColumn8.NullText = ""
        Me.DataGridTextBoxColumn8.ReadOnly = True
        Me.DataGridTextBoxColumn8.Width = 75
        '
        'DataGridTextBoxColumn3
        '
        Me.DataGridTextBoxColumn3.Format = ""
        Me.DataGridTextBoxColumn3.FormatInfo = Nothing
        Me.DataGridTextBoxColumn3.HeaderText = "原值(元)"
        Me.DataGridTextBoxColumn3.MappingName = "original_value"
        Me.DataGridTextBoxColumn3.NullText = "0"
        Me.DataGridTextBoxColumn3.ReadOnly = True
        Me.DataGridTextBoxColumn3.Width = 75
        '
        'DataGridTextBoxColumn10
        '
        Me.DataGridTextBoxColumn10.Format = ""
        Me.DataGridTextBoxColumn10.FormatInfo = Nothing
        Me.DataGridTextBoxColumn10.HeaderText = "评估值(元)"
        Me.DataGridTextBoxColumn10.MappingName = "evaluate_value"
        Me.DataGridTextBoxColumn10.NullText = "0"
        Me.DataGridTextBoxColumn10.ReadOnly = True
        Me.DataGridTextBoxColumn10.Width = 75
        '
        'DataGridTextBoxColumn11
        '
        Me.DataGridTextBoxColumn11.Format = ""
        Me.DataGridTextBoxColumn11.FormatInfo = Nothing
        Me.DataGridTextBoxColumn11.HeaderText = "担保额(元)"
        Me.DataGridTextBoxColumn11.MappingName = "guarantee_value"
        Me.DataGridTextBoxColumn11.NullText = "0"
        Me.DataGridTextBoxColumn11.ReadOnly = True
        Me.DataGridTextBoxColumn11.Width = 75
        '
        'DataGridTextBoxColumn9
        '
        Me.DataGridTextBoxColumn9.Format = ""
        Me.DataGridTextBoxColumn9.FormatInfo = Nothing
        Me.DataGridTextBoxColumn9.HeaderText = "备注"
        Me.DataGridTextBoxColumn9.MappingName = "remark"
        Me.DataGridTextBoxColumn9.NullText = ""
        Me.DataGridTextBoxColumn9.ReadOnly = True
        Me.DataGridTextBoxColumn9.Width = 75
        '
        'DataGridTextBoxColumn14
        '
        Me.DataGridTextBoxColumn14.Format = ""
        Me.DataGridTextBoxColumn14.FormatInfo = Nothing
        Me.DataGridTextBoxColumn14.HeaderText = "状态"
        Me.DataGridTextBoxColumn14.MappingName = "status"
        Me.DataGridTextBoxColumn14.NullText = ""
        Me.DataGridTextBoxColumn14.Width = 75
        '
        'Label14
        '
        Me.Label14.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(472, 323)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(60, 14)
        Me.Label14.TabIndex = 32
        Me.Label14.Text = "项目经理A"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtManagerA
        '
        Me.txtManagerA.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.txtManagerA.Location = New System.Drawing.Point(544, 320)
        Me.txtManagerA.Name = "txtManagerA"
        Me.txtManagerA.ReadOnly = True
        Me.txtManagerA.Size = New System.Drawing.Size(144, 21)
        Me.txtManagerA.TabIndex = 33
        Me.txtManagerA.Text = ""
        '
        'dtpBooking
        '
        Me.dtpBooking.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.dtpBooking.Enabled = False
        Me.dtpBooking.Location = New System.Drawing.Point(88, 349)
        Me.dtpBooking.Name = "dtpBooking"
        Me.dtpBooking.Size = New System.Drawing.Size(128, 21)
        Me.dtpBooking.TabIndex = 34
        '
        'Label15
        '
        Me.Label15.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(8, 352)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(79, 14)
        Me.Label15.TabIndex = 35
        Me.Label15.Text = "预约评估时间"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'FAffirmEvaluateDate
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.ClientSize = New System.Drawing.Size(698, 527)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label15, Me.Label14, Me.Label8, Me.Label12, Me.Label13, Me.txtName, Me.btnExit, Me.GroupBox4, Me.GroupBox3, Me.GroupBox2, Me.dtpBooking, Me.txtManagerA, Me.btnSave, Me.btnCommit, Me.btnViewPic, Me.dtpEvaluateAffirmDate, Me.corp_contact_phone, Me.corp_contact_person})
        Me.Name = "FAffirmEvaluateDate"
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.dgOppGuarantee, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private StrProjectCode, StrPhase As String
    Private StrSerial As String '序列号

    Private DsOpp As DataSet
    Private DvDetail As DataView

    Private CurrentIndex As Int16

    Private HasChanged As Boolean

    Private Const OPP As String = "opposite_guarantee"
    Private Const DETAIL As String = "opposite_guarantee_detail"

    Private Sub FAffirmEvaluateDate_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.strProjectCode = Me.getProjectCode()
        Me.StrPhase = Me.getPhase(Me.StrProjectCode)

        Me.txtCode.Text = Me.strProjectCode
        Me.txtName.Text = Me.getCorporationName()

        Call Me.getProjectCorporation(Me.StrProjectCode, Me.StrPhase) '获取联系人,联系电话
        Me.txtManagerA.Text = Me.getManagerA(Me.StrProjectCode)  '设置项目经理a
        getOppGuaranteeType() '设置类型combobox
        getOppGuaranteeMeasure() '设置措施combobox

        Me.DsOpp = Me.getProjectOpp()

        If Not Me.DsOpp Is Nothing Then
            Me.dgOppGuarantee.DataSource = Me.DsOpp.Tables(OPP)
            setDgAttribute(Me.dgOppGuarantee)
        End If
        If Me.DsOpp.Tables(OPP).Rows.Count > 0 And Not Me.DsOpp.Tables(DETAIL) Is Nothing Then
            If Me.dgOppGuarantee.CurrentRowIndex >= 0 Then
                Me.CurrentIndex = Me.dgOppGuarantee.CurrentRowIndex
            Else
                Me.CurrentIndex = 0
            End If

            Me.setValue(Me.CurrentIndex)

            Me.StrSerial = Me.DsOpp.Tables(OPP).Rows(0).Item("serial_num")
            Me.DvDetail = Me.DsOpp.Tables(DETAIL).DefaultView
            Me.DvDetail.AllowNew = False
            Me.DvDetail.RowFilter = "serial_num='" + Me.StrSerial + "'"
            If Not Me.DvDetail Is Nothing Then
                Me.dgDetail.DataSource = Me.DvDetail
                Me.addStyle(Me.DsOpp)
            End If
            setButton()
        End If

        AddHandler dtpEvaluateAffirmDate.ValueChanged, AddressOf dtpEvaluateAffirmDate_ValueChanged

    End Sub
    '设置控件对应记录的显示值
    Private Sub setValue(ByVal index)

        'Dim index As Integer
        Dim dt As DataTable
        'index = Me.dgOppGuarantee.CurrentRowIndex
        dt = CType(Me.dgOppGuarantee.DataSource, DataTable)
        If index < 0 Then
            Exit Sub
        End If
        Try
            With dt.Rows(index)
                If Not .Item("original_value") Is System.DBNull.Value Then
                    Me.txtOriginal.Text = .Item("original_value")
                Else
                    Me.txtOriginal.Text = ""
                End If
                If Not .Item("evaluate_value") Is System.DBNull.Value Then
                    Me.txtEvaluate.Text = .Item("evaluate_value")
                Else
                    Me.txtEvaluate.Text = ""
                End If
                If Not .Item("guarantee_value") Is System.DBNull.Value Then
                    Me.txtGuarantee.Text = .Item("guarantee_value")
                Else
                    Me.txtGuarantee.Text = ""
                End If
                If Not .Item("evaluate_date") Is System.DBNull.Value Then
                    Me.dtpBooking.Value = .Item("evaluate_date")
                Else
                    Try
                        Me.dtpBooking.Value = gWs.GetSysTime()
                    Catch ex As Exception
                        ExceptionHandler.ShowMessageBox(ex)
                    End Try
                End If
                If Not .Item("contract_no") Is System.DBNull.Value Then
                    Me.txtContract.Text = .Item("contract_no")
                Else
                    Me.txtContract.Text = ""
                End If
                If Not .Item("remark") Is System.DBNull.Value Then
                    Me.txtRemark.Text = .Item("remark")
                Else
                    Me.txtRemark.Text = ""
                End If
                If Not .Item("guaranty_type") Is System.DBNull.Value Then
                    Me.cmbType.Text = getGuaranteeType(.Item("guaranty_type"))
                End If
                If Not .Item("opposite_guarantee_form") Is System.DBNull.Value Then
                    Me.cmbMeasure.Text = getGuaranteeForm(.Item("opposite_guarantee_form"))
                End If
                If Not .Item("affirm_evaluate_date") Is System.DBNull.Value Then
                    Me.dtpEvaluateAffirmDate.Value = .Item("affirm_evaluate_date")
                End If
            End With
        Catch ex As Exception
            'MsgBox(ex.Message)
            ExceptionHandler.ShowMessageBox(ex)
        End Try

    End Sub
    '获取企业登记了的反担保措施
    Private Function getProjectOpp() As DataSet

        Dim strSql, strSqlDetail As String

        Dim dsTemp As DataSet

        strSql = "{project_code ='" & StrProjectCode & "'}"
        strSqlDetail = "{project_code ='" & StrProjectCode & "'}"
        Try

            dsTemp = gWs.GetGuarantyInfo(strSql, strSqlDetail)

            If dsTemp Is Nothing Then Return Nothing

            dsTemp.Tables("opposite_guarantee").Columns.Add("type_name", GetType(String))
            dsTemp.Tables("opposite_guarantee_detail").Columns.Add("item_name", GetType(String))

            Dim j As Integer
            j = dsTemp.Tables("opposite_guarantee").Rows.Count

            Dim dr As DataRow
            Dim strItemType As String
            Dim dsItemType As DataSet

            For j = 0 To j - 1
                dr = dsTemp.Tables("opposite_guarantee").Rows(j)
                Try
                    If Not dr.Item("guaranty_type") Is System.DBNull.Value Then
                        strItemType = dr.Item("guaranty_type")
                        dsItemType = getItemTypeDataSet(strItemType)
                        'oppGuarantee显示的类型名映射为ItemType中的typeName
                        dr.Item("type_name") = dsItemType.Tables("TItemType").Select("item_type='" & strItemType & "'")(0).Item("type_name")
                    End If
                    If dr.Item("affirm_evaluate_date") Is System.DBNull.Value Then
                        dr.Item("affirm_evaluate_date") = dr.Item("evaluate_date")
                    End If
                Catch ex As Exception
                    'MsgBox("获取反担保物类型出错!" & ex.Message, MsgBoxStyle.Critical, "提示")
                    ExceptionHandler.ShowMessageBox(ex)
                End Try
            Next

            j = dsTemp.Tables(DETAIL).Rows.Count
            For j = 0 To j - 1
                dr = dsTemp.Tables(DETAIL).Rows(j)
                Try
                    If Not dr.Item("item_type") Is System.DBNull.Value And Not dr.Item("item_code") Is System.DBNull.Value Then
                        dr.Item("item_name") = Me.getItemName(dr.Item("item_type"), dr.Item("item_code"))
                    End If
                Catch ex As Exception

                End Try
            Next

            'dsTemp.AcceptChanges()

            Return dsTemp

        Catch ex As Exception
            ExceptionHandler.ShowMessageBox(ex)
            Return Nothing
        End Try

    End Function

    '获得项目企业表信息,并设置“企业联系人”和“联系电话”的初始值
    Private Sub getProjectCorporation(ByVal projectCode As String, ByVal phase As String)
        Try
            Dim strSql As String
            Dim dsProjectCorporation As DataSet
            strSql = "{project_code='" & projectCode & "' and phase='" & phase & "' and corporation_code='" & projectCode.Substring(0, 5) & "'}"
            'dsProjectCorporation = gWs.FetchProjectCorporation(strSql)
            dsProjectCorporation = gWs.FetchProjectCorporation(Me.getProjectCode(), "%", "1", Me.getPhase(getProjectCode()))
            If Not dsProjectCorporation Is Nothing Then
                If dsProjectCorporation.Tables(0).Rows.Count > 0 Then
                    If Not dsProjectCorporation.Tables(0).Rows(0).Item("corp_contact_person") Is System.DBNull.Value Then
                        Me.corp_contact_person.Text = dsProjectCorporation.Tables(0).Rows(0).Item("corp_contact_person")
                    End If
                    If Not dsProjectCorporation.Tables(0).Rows(0).Item("corp_contact_phone") Is System.DBNull.Value Then
                        Me.corp_contact_phone.Text = dsProjectCorporation.Tables(0).Rows(0).Item("corp_contact_phone")
                    End If
                    'hasProCorpChanged = False
                End If
            End If
        Catch ex As Exception
            ExceptionHandler.ShowMessageBox(ex)
        End Try
    End Sub
    '列表中鼠标按下事件
    Private Sub dgOppGuarantee_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgOppGuarantee.MouseUp
        If Not Me.DsOpp Is Nothing AndAlso Me.DsOpp.Tables(OPP).Rows.Count > 0 Then
       
            Me.CurrentIndex = Me.dgOppGuarantee.CurrentRowIndex
            Me.setValue(Me.CurrentIndex)

            Me.StrSerial = Me.DsOpp.Tables(OPP).Rows(Me.CurrentIndex).Item("guaranty_type")

            Me.DvDetail.RowFilter = "item_type='" + Me.StrSerial + "'"
            setButton()
        End If
    End Sub
    '设置按钮的Enabled属性
    Private Sub setButton()
        If Me.DsOpp.Tables(OPP).Rows(Me.CurrentIndex).Item("file_relation_num") Is System.DBNull.Value Then
            Me.btnViewPic.Enabled = False
        Else
            Me.btnViewPic.Enabled = True
        End If
        'qxd add 2004-8-10
        If Me.DsOpp.Tables(OPP).Rows(Me.CurrentIndex).Item("evaluate_person") Is System.DBNull.Value Then
            Me.dtpEvaluateAffirmDate.Enabled = False
            Me.btnSave.Enabled = False
        ElseIf Me.DsOpp.Tables(OPP).Rows(Me.CurrentIndex).Item("evaluate_person") = UserName Then
            Me.dtpEvaluateAffirmDate.Enabled = True
            Me.btnSave.Enabled = True
        Else
            Me.dtpEvaluateAffirmDate.Enabled = False
            Me.btnSave.Enabled = False
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.DsOpp.HasChanges Then
            Dim strResult As String
            strResult = gWs.UpdateGuaranty(Me.DsOpp)
            If strResult = "1" Then
                SWDialogBox.MessageBox.Show("$SaveSucceed")
                Me.DsOpp.AcceptChanges()
                Me.HasChanged = False
            Else
                SWDialogBox.MessageBox.Show("*999", "", strResult, "")
            End If
        End If
    End Sub

    Private Sub dtpEvaluateAffirmDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not Me.DsOpp Is Nothing AndAlso Me.DsOpp.Tables(OPP).Rows.Count > 0 Then
            If Not Me.dtpEvaluateAffirmDate.Value = Me.DsOpp.Tables(OPP).Rows(Me.CurrentIndex).Item("affirm_evaluate_date") Then
                Me.HasChanged = True
                'qxd add 2004-8-10
                Dim count, i As Integer
                count = Me.DsOpp.Tables(OPP).Rows.Count
                If count > 0 Then
                    For i = 0 To count - 1
                        If Not Me.DsOpp.Tables(OPP).Rows(i).Item("evaluate_person") Is System.DBNull.Value Then
                            If Me.DsOpp.Tables(OPP).Rows(i).Item("evaluate_person") = UserName Then
                                Me.DsOpp.Tables(OPP).Rows(i).Item("affirm_evaluate_date") = Me.dtpEvaluateAffirmDate.Value
                            End If
                        End If
                    Next
                End If
                'Me.DsOpp.Tables(OPP).Rows(Me.CurrentIndex).Item("affirm_evaluate_date") = Me.dtpEvaluateAffirmDate.Value
            End If
        End If
    End Sub

    '浏览
    Private Sub btnViewPic_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnViewPic.Click
        Dim ds As DataSet
        Dim strCode As String
        Dim strRelationNum As String
        Dim strItemType As String

        Dim index As Integer
        index = Me.dgOppGuarantee.CurrentRowIndex
        If index < 0 Then
            Exit Sub
        End If

        ds = Me.DsOpp
        If Not ds.Tables("opposite_guarantee").Rows(index).Item("project_code") Is System.DBNull.Value Then
            strCode = ds.Tables("opposite_guarantee").Rows(index).Item("project_code")
        End If
        If Not ds.Tables("opposite_guarantee").Rows(index).Item("guaranty_type") Is System.DBNull.Value Then
            strItemType = ds.Tables("opposite_guarantee").Rows(index).Item("guaranty_type")
        End If
        If Not ds.Tables("opposite_guarantee").Rows(index).Item("file_relation_num") Is System.DBNull.Value Then
            strRelationNum = ds.Tables("opposite_guarantee").Rows(index).Item("file_relation_num")
            Me.Cursor = Cursors.WaitCursor
            Dim docClass As New frmDocumentClass("")
            'docClass.openFileList(StrProjectCode, strItemType, "", "\Document\Image\")
            docClass.openFileList(StrProjectCode, "2%", "", "\Document\Image\")
            Me.Cursor = Cursors.Default
        Else
            'MsgBox("没有相关的图片!", MsgBoxStyle.Exclamation, "提示")
            SWDialogBox.MessageBox.Show("$1005", "图片")
        End If
    End Sub

    Private Sub btnCommit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCommit.Click
        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.isCommitConfirm Then
                If MsgBoxResult.Yes = SWDialogBox.MessageBox.Show("?1000", "提交") Then
                    commit()
                End If
            Else

            End If
        Catch ex As Exception
            ExceptionHandler.ShowMessageBox(ex)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub
    Private Sub commit()

        Dim strResult As String
        If Me.DsOpp.HasChanges Then
            strResult = gWs.UpdateGuaranty(Me.DsOpp)
            If Not strResult = "1" Then
                SWDialogBox.MessageBox.Show("*999", "", strResult, "")
            End If
        End If

        Try
            strResult = gWs.finishedTask(Me.getWorkFlowID(), Me.StrProjectCode, Me.getTaskID(), "", Me.getUser(), 1)
        Catch ex As Exception
            ExceptionHandler.ShowMessageBox(ex)
        End Try
        If strResult = "1" Then
            'MsgBox("提交成功！", MsgBoxStyle.Information, "提示")
            SWDialogBox.MessageBox.Show("$SubmitSucceed")
            raiseCommitSucceed()
            Me.Close()
        Else
            'MsgBox("提交失败！" & strResult, MsgBoxStyle.Critical, "提示")
            SWDialogBox.MessageBox.Show("*999", "", strResult, "")
        End If
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click

        Dim strMesg As Integer

        If Me.HasChanged Then
            'strMesg = MsgBox("数据已修改，是否保存？", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNoCancel, "提示")
            strMesg = SWDialogBox.MessageBox.Show("?1003")
        Else
            Me.Close()
            Exit Sub
        End If

        If MsgBoxResult.Yes = strMesg Then
            Call btnSave_Click(sender, e)
            Me.Close()
        ElseIf MsgBoxResult.No = strMesg Then
            Me.Close()
        Else
            Exit Sub
        End If
    End Sub
    Private Function getItemName(ByVal strItemType As String, ByVal strItemCode As String) As String
        Try
            Return gWs.GetItem(strItemCode, strItemType).Tables(0).Rows(0).Item("item_name")
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Private Function getManagerA(ByVal strProjectCode) As String
        Try
            Dim strSql As String = "{projectcode='" & strProjectCode & "'}"
            Return gWs.GetProjectInfoEx(strSql).Tables(0).Rows(0).Item("Manager_a")
        Catch ex As Exception
            Return ""
        End Try
    End Function

End Class
