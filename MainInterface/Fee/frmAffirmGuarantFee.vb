

Public Class frmAffirmGuarantFee
    Inherits frmAffirmFeeBase

#Region " Windows ������������ɵĴ��� "

    Public Sub New()
        MyBase.New()

        '�õ����� Windows ���������������ġ�
        InitializeComponent()

        '�� InitializeComponent() ����֮������κγ�ʼ��

    End Sub

    '������д��������������б�
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Windows ����������������
    Private components As System.ComponentModel.IContainer

    'ע�⣺���¹����� Windows ����������������
    '����ʹ�� Windows ����������޸Ĵ˹��̡�
    '��Ҫʹ�ô���༭���޸�����
    Friend WithEvents txtManager As System.Windows.Forms.TextBox
    Friend WithEvents lblManager As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtManager = New System.Windows.Forms.TextBox()
        Me.lblManager = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label8
        '
        Me.Label8.Text = "���յ�����(Ԫ)"
        Me.Label8.Visible = True
        '
        'txtIncome
        '
        Me.txtIncome.Visible = True
        '
        'txtPayout
        '
        Me.txtPayout.Visible = True
        '
        'lblFeeType
        '
        Me.lblFeeType.Text = "Ӧ�յ�����(Ԫ)"
        Me.lblFeeType.Visible = True
        '
        'txtCorName
        '
        Me.txtCorName.Visible = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtManager, Me.lblManager})
        Me.GroupBox1.Visible = True
        '
        'txtProjectCode
        '
        Me.txtProjectCode.Visible = True
        '
        'Label6
        '
        Me.Label6.Visible = True
        '
        'Label5
        '
        Me.Label5.Visible = True
        '
        'btnExit
        '
        Me.btnExit.Visible = True
        '
        'txtManager
        '
        Me.txtManager.BackColor = System.Drawing.SystemColors.Window
        Me.txtManager.Enabled = False
        Me.txtManager.Location = New System.Drawing.Point(256, 40)
        Me.txtManager.Name = "txtManager"
        Me.txtManager.Size = New System.Drawing.Size(80, 21)
        Me.txtManager.TabIndex = 37
        Me.txtManager.Text = ""
        '
        'lblManager
        '
        Me.lblManager.Location = New System.Drawing.Point(188, 42)
        Me.lblManager.Name = "lblManager"
        Me.lblManager.Size = New System.Drawing.Size(64, 17)
        Me.lblManager.TabIndex = 36
        Me.lblManager.Text = "��Ŀ����A"
        Me.lblManager.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'frmAffirmGuarantFee
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.ClientSize = New System.Drawing.Size(538, 335)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnExit, Me.GroupBox1})
        Me.Name = "frmAffirmGuarantFee"
        Me.Text = "ȷ�ϵ�������ȡ"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frm_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Try
            Me.Cursor = Cursors.WaitCursor
            ProjectCode = MyBase.getProjectCode
            CorporationName = MyBase.getCorporationName
            TaskID = MyBase.getTaskID
            WorkFlowID = MyBase.getWorkFlowID

            IsGuarantyFee = True
            SystemDate = gWs.GetSysTime
            GetDataSource("31", "002")
            txtPayout.Text = Payout.ToString("c", CGFormatInfo)
            txtincome.Text = Income.ToString("c", CGFormatInfo)

            Dim dsMan As DataSet = New DataSet()
            dsMan = gWs.GetProjectInfoEx("{ProjectCode LIKE '" & ProjectCode & "'}")
            txtManager.DataBindings.Add("Text", dsMan, "ViewProject.24")
        Catch ex As Exception
            SWDialogBox.MessageBox.Show("*999", ex.Source, ex.Message, ex.StackTrace)
            'MessageBox.Show("��������" & vbCr & ex.Message, "��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub
End Class
