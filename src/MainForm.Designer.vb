'
' 日付: 2016/06/18
'
Partial Class MainForm
	Inherits System.Windows.Forms.Form
	
	''' <summary>
	''' Designer variable used to keep track of non-visual components.
	''' </summary>
	Private components As System.ComponentModel.IContainer
	
	''' <summary>
	''' Disposes resources used by the form.
	''' </summary>
	''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
	Protected Overrides Sub Dispose(ByVal disposing As Boolean)
		If disposing Then
			If components IsNot Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(disposing)
	End Sub
	
	''' <summary>
	''' This method is required for Windows Forms designer support.
	''' Do not change the method contents inside the source code editor. The Forms designer might
	''' not be able to load this method if it was changed manually.
	''' </summary>
	Private Sub InitializeComponent()
	  Me.tabRoot = New System.Windows.Forms.TabControl()
	  Me.tPagePersonal = New System.Windows.Forms.TabPage()
	  Me.tabInPersonalTab = New System.Windows.Forms.TabControl()
	  Me.tPageMonth10InPersonal = New System.Windows.Forms.TabPage()
	  Me.gridMonth10InPersonal = New System.Windows.Forms.DataGridView()
	  Me.cboxUserName = New System.Windows.Forms.ComboBox()
	  Me.tPageDate = New System.Windows.Forms.TabPage()
	  Me.tabInDateTab = New System.Windows.Forms.TabControl()
	  Me.tPageDailyInDate = New System.Windows.Forms.TabPage()
	  Me.gridDailyInDate = New System.Windows.Forms.DataGridView()
	  Me.dateTimePickerInDatePage = New System.Windows.Forms.DateTimePicker()
	  Me.tPageWeeklyInDate = New System.Windows.Forms.TabPage()
	  Me.gridWeeklyInDate = New System.Windows.Forms.DataGridView()
	  Me.cboxWeekly = New System.Windows.Forms.ComboBox()
	  Me.tPageMonthlyInDate = New System.Windows.Forms.TabPage()
	  Me.gridMonthlyInDate = New System.Windows.Forms.DataGridView()
	  Me.cboxMonthly = New System.Windows.Forms.ComboBox()
	  Me.tPageYearlyInDate = New System.Windows.Forms.TabPage()
	  Me.gridTallyInDate = New System.Windows.Forms.DataGridView()
	  Me.tPageTally = New System.Windows.Forms.TabPage()
	  Me.tabInTallyTab = New System.Windows.Forms.TabControl()
	  Me.tPageDailyInTally = New System.Windows.Forms.TabPage()
	  Me.dataGridView1 = New System.Windows.Forms.DataGridView()
	  Me.cboxTallyMonthly = New System.Windows.Forms.ComboBox()
	  Me.tPageWeeklyInTally = New System.Windows.Forms.TabPage()
	  Me.dataGridView2 = New System.Windows.Forms.DataGridView()
	  Me.tPageMonthlyInTally = New System.Windows.Forms.TabPage()
	  Me.dataGridView3 = New System.Windows.Forms.DataGridView()
	  Me.btnClose = New System.Windows.Forms.Button()
	  Me.btnReload = New System.Windows.Forms.Button()
	  Me.btnOutputCsv = New System.Windows.Forms.Button()
	  Me.chkBoxExcludeData = New System.Windows.Forms.CheckBox()
	  Me.tabRoot.SuspendLayout
	  Me.tPagePersonal.SuspendLayout
	  Me.tabInPersonalTab.SuspendLayout
	  Me.tPageMonth10InPersonal.SuspendLayout
	  CType(Me.gridMonth10InPersonal,System.ComponentModel.ISupportInitialize).BeginInit
	  Me.tPageDate.SuspendLayout
	  Me.tabInDateTab.SuspendLayout
	  Me.tPageDailyInDate.SuspendLayout
	  CType(Me.gridDailyInDate,System.ComponentModel.ISupportInitialize).BeginInit
	  Me.tPageWeeklyInDate.SuspendLayout
	  CType(Me.gridWeeklyInDate,System.ComponentModel.ISupportInitialize).BeginInit
	  Me.tPageMonthlyInDate.SuspendLayout
	  CType(Me.gridMonthlyInDate,System.ComponentModel.ISupportInitialize).BeginInit
	  Me.tPageYearlyInDate.SuspendLayout
	  CType(Me.gridTallyInDate,System.ComponentModel.ISupportInitialize).BeginInit
	  Me.tPageTally.SuspendLayout
	  Me.tabInTallyTab.SuspendLayout
	  Me.tPageDailyInTally.SuspendLayout
	  CType(Me.dataGridView1,System.ComponentModel.ISupportInitialize).BeginInit
	  Me.tPageWeeklyInTally.SuspendLayout
	  CType(Me.dataGridView2,System.ComponentModel.ISupportInitialize).BeginInit
	  Me.tPageMonthlyInTally.SuspendLayout
	  CType(Me.dataGridView3,System.ComponentModel.ISupportInitialize).BeginInit
	  Me.SuspendLayout
	  '
	  'tabRoot
	  '
	  Me.tabRoot.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom)  _
	  	  	  Or System.Windows.Forms.AnchorStyles.Left)  _
	  	  	  Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
	  Me.tabRoot.Controls.Add(Me.tPagePersonal)
	  Me.tabRoot.Controls.Add(Me.tPageDate)
	  Me.tabRoot.Controls.Add(Me.tPageTally)
	  Me.tabRoot.Location = New System.Drawing.Point(0, 0)
	  Me.tabRoot.Margin = New System.Windows.Forms.Padding(0)
	  Me.tabRoot.Name = "tabRoot"
	  Me.tabRoot.SelectedIndex = 0
	  Me.tabRoot.Size = New System.Drawing.Size(771, 355)
	  Me.tabRoot.TabIndex = 0
	  AddHandler Me.tabRoot.SelectedIndexChanged, AddressOf Me.TabRoot_SelectedIndexChanged
	  '
	  'tPagePersonal
	  '
	  Me.tPagePersonal.BackColor = System.Drawing.SystemColors.Control
	  Me.tPagePersonal.Controls.Add(Me.tabInPersonalTab)
	  Me.tPagePersonal.Controls.Add(Me.cboxUserName)
	  Me.tPagePersonal.Location = New System.Drawing.Point(4, 22)
	  Me.tPagePersonal.Margin = New System.Windows.Forms.Padding(0)
	  Me.tPagePersonal.Name = "tPagePersonal"
	  Me.tPagePersonal.Padding = New System.Windows.Forms.Padding(3)
	  Me.tPagePersonal.Size = New System.Drawing.Size(763, 329)
	  Me.tPagePersonal.TabIndex = 0
	  Me.tPagePersonal.Text = "個人データ"
	  '
	  'tabInPersonalTab
	  '
	  Me.tabInPersonalTab.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom)  _
	  	  	  Or System.Windows.Forms.AnchorStyles.Left)  _
	  	  	  Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
	  Me.tabInPersonalTab.Controls.Add(Me.tPageMonth10InPersonal)
	  Me.tabInPersonalTab.Location = New System.Drawing.Point(8, 29)
	  Me.tabInPersonalTab.Name = "tabInPersonalTab"
	  Me.tabInPersonalTab.SelectedIndex = 0
	  Me.tabInPersonalTab.Size = New System.Drawing.Size(749, 294)
	  Me.tabInPersonalTab.TabIndex = 1
	  AddHandler Me.tabInPersonalTab.SelectedIndexChanged, AddressOf Me.TabInPersonalTab_SelectedIndexChanged
	  '
	  'tPageMonth10InPersonal
	  '
	  Me.tPageMonth10InPersonal.BackColor = System.Drawing.SystemColors.Control
	  Me.tPageMonth10InPersonal.Controls.Add(Me.gridMonth10InPersonal)
	  Me.tPageMonth10InPersonal.Location = New System.Drawing.Point(4, 22)
	  Me.tPageMonth10InPersonal.Name = "tPageMonth10InPersonal"
	  Me.tPageMonth10InPersonal.Padding = New System.Windows.Forms.Padding(3)
	  Me.tPageMonth10InPersonal.Size = New System.Drawing.Size(741, 268)
	  Me.tPageMonth10InPersonal.TabIndex = 0
	  Me.tPageMonth10InPersonal.Text = "page1"
	  '
	  'gridMonth10InPersonal
	  '
	  Me.gridMonth10InPersonal.AllowUserToAddRows = false
	  Me.gridMonth10InPersonal.AllowUserToDeleteRows = false
	  Me.gridMonth10InPersonal.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
	  Me.gridMonth10InPersonal.Dock = System.Windows.Forms.DockStyle.Fill
	  Me.gridMonth10InPersonal.Location = New System.Drawing.Point(3, 3)
	  Me.gridMonth10InPersonal.Name = "gridMonth10InPersonal"
	  Me.gridMonth10InPersonal.ReadOnly = true
	  Me.gridMonth10InPersonal.RowTemplate.Height = 21
	  Me.gridMonth10InPersonal.Size = New System.Drawing.Size(735, 262)
	  Me.gridMonth10InPersonal.TabIndex = 0
	  '
	  'cboxUserName
	  '
	  Me.cboxUserName.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
	  Me.cboxUserName.FormattingEnabled = true
	  Me.cboxUserName.Location = New System.Drawing.Point(557, 6)
	  Me.cboxUserName.Name = "cboxUserName"
	  Me.cboxUserName.Size = New System.Drawing.Size(200, 20)
	  Me.cboxUserName.TabIndex = 0
	  Me.cboxUserName.Text = "参照するユーザを選んでください"
	  AddHandler Me.cboxUserName.SelectedIndexChanged, AddressOf Me.CboxUserNameSelectedIndexChanged
	  '
	  'tPageDate
	  '
	  Me.tPageDate.BackColor = System.Drawing.SystemColors.Control
	  Me.tPageDate.Controls.Add(Me.tabInDateTab)
	  Me.tPageDate.Location = New System.Drawing.Point(4, 22)
	  Me.tPageDate.Name = "tPageDate"
	  Me.tPageDate.Padding = New System.Windows.Forms.Padding(3)
	  Me.tPageDate.Size = New System.Drawing.Size(763, 329)
	  Me.tPageDate.TabIndex = 1
	  Me.tPageDate.Text = "日付データ"
	  '
	  'tabInDateTab
	  '
	  Me.tabInDateTab.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom)  _
	  	  	  Or System.Windows.Forms.AnchorStyles.Left)  _
	  	  	  Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
	  Me.tabInDateTab.Controls.Add(Me.tPageDailyInDate)
	  Me.tabInDateTab.Controls.Add(Me.tPageWeeklyInDate)
	  Me.tabInDateTab.Controls.Add(Me.tPageMonthlyInDate)
	  Me.tabInDateTab.Controls.Add(Me.tPageYearlyInDate)
	  Me.tabInDateTab.Location = New System.Drawing.Point(8, 29)
	  Me.tabInDateTab.Name = "tabInDateTab"
	  Me.tabInDateTab.SelectedIndex = 0
	  Me.tabInDateTab.Size = New System.Drawing.Size(749, 294)
	  Me.tabInDateTab.TabIndex = 2
	  AddHandler Me.tabInDateTab.SelectedIndexChanged, AddressOf Me.TabInDateTab_SelectedIndexChanged
	  '
	  'tPageDailyInDate
	  '
	  Me.tPageDailyInDate.BackColor = System.Drawing.SystemColors.Control
	  Me.tPageDailyInDate.Controls.Add(Me.gridDailyInDate)
	  Me.tPageDailyInDate.Controls.Add(Me.dateTimePickerInDatePage)
	  Me.tPageDailyInDate.Location = New System.Drawing.Point(4, 22)
	  Me.tPageDailyInDate.Name = "tPageDailyInDate"
	  Me.tPageDailyInDate.Padding = New System.Windows.Forms.Padding(3)
	  Me.tPageDailyInDate.Size = New System.Drawing.Size(741, 268)
	  Me.tPageDailyInDate.TabIndex = 0
	  Me.tPageDailyInDate.Text = "日"
	  '
	  'gridDailyInDate
	  '
	  Me.gridDailyInDate.AllowUserToAddRows = false
	  Me.gridDailyInDate.AllowUserToDeleteRows = false
	  Me.gridDailyInDate.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom)  _
	  	  	  Or System.Windows.Forms.AnchorStyles.Left)  _
	  	  	  Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
	  Me.gridDailyInDate.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
	  Me.gridDailyInDate.Location = New System.Drawing.Point(3, 31)
	  Me.gridDailyInDate.Name = "gridDailyInDate"
	  Me.gridDailyInDate.ReadOnly = true
	  Me.gridDailyInDate.RowTemplate.Height = 21
	  Me.gridDailyInDate.Size = New System.Drawing.Size(735, 234)
	  Me.gridDailyInDate.TabIndex = 3
	  '
	  'dateTimePickerInDatePage
	  '
	  Me.dateTimePickerInDatePage.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
	  Me.dateTimePickerInDatePage.Location = New System.Drawing.Point(561, 6)
	  Me.dateTimePickerInDatePage.Name = "dateTimePickerInDatePage"
	  Me.dateTimePickerInDatePage.Size = New System.Drawing.Size(160, 19)
	  Me.dateTimePickerInDatePage.TabIndex = 2
	  AddHandler Me.dateTimePickerInDatePage.ValueChanged, AddressOf Me.DateTimePickerInDatePageValueChanged
	  '
	  'tPageWeeklyInDate
	  '
	  Me.tPageWeeklyInDate.BackColor = System.Drawing.SystemColors.Control
	  Me.tPageWeeklyInDate.Controls.Add(Me.gridWeeklyInDate)
	  Me.tPageWeeklyInDate.Controls.Add(Me.cboxWeekly)
	  Me.tPageWeeklyInDate.Location = New System.Drawing.Point(4, 22)
	  Me.tPageWeeklyInDate.Name = "tPageWeeklyInDate"
	  Me.tPageWeeklyInDate.Padding = New System.Windows.Forms.Padding(3)
	  Me.tPageWeeklyInDate.Size = New System.Drawing.Size(741, 268)
	  Me.tPageWeeklyInDate.TabIndex = 1
	  Me.tPageWeeklyInDate.Text = "週"
	  '
	  'gridWeeklyInDate
	  '
	  Me.gridWeeklyInDate.AllowUserToAddRows = false
	  Me.gridWeeklyInDate.AllowUserToDeleteRows = false
	  Me.gridWeeklyInDate.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom)  _
	  	  	  Or System.Windows.Forms.AnchorStyles.Left)  _
	  	  	  Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
	  Me.gridWeeklyInDate.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
	  Me.gridWeeklyInDate.Location = New System.Drawing.Point(3, 31)
	  Me.gridWeeklyInDate.Name = "gridWeeklyInDate"
	  Me.gridWeeklyInDate.ReadOnly = true
	  Me.gridWeeklyInDate.RowTemplate.Height = 21
	  Me.gridWeeklyInDate.Size = New System.Drawing.Size(735, 234)
	  Me.gridWeeklyInDate.TabIndex = 4
	  '
	  'cboxWeekly
	  '
	  Me.cboxWeekly.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
	  Me.cboxWeekly.FormattingEnabled = true
	  Me.cboxWeekly.Location = New System.Drawing.Point(560, 6)
	  Me.cboxWeekly.Name = "cboxWeekly"
	  Me.cboxWeekly.Size = New System.Drawing.Size(160, 20)
	  Me.cboxWeekly.TabIndex = 1
	  Me.cboxWeekly.Text = "参照する週を選んでください"
	  AddHandler Me.cboxWeekly.SelectedIndexChanged, AddressOf Me.CboxWeeklySelectedIndexChanged
	  '
	  'tPageMonthlyInDate
	  '
	  Me.tPageMonthlyInDate.BackColor = System.Drawing.SystemColors.Control
	  Me.tPageMonthlyInDate.Controls.Add(Me.gridMonthlyInDate)
	  Me.tPageMonthlyInDate.Controls.Add(Me.cboxMonthly)
	  Me.tPageMonthlyInDate.Location = New System.Drawing.Point(4, 22)
	  Me.tPageMonthlyInDate.Name = "tPageMonthlyInDate"
	  Me.tPageMonthlyInDate.Padding = New System.Windows.Forms.Padding(3)
	  Me.tPageMonthlyInDate.Size = New System.Drawing.Size(741, 268)
	  Me.tPageMonthlyInDate.TabIndex = 2
	  Me.tPageMonthlyInDate.Text = "月"
	  '
	  'gridMonthlyInDate
	  '
	  Me.gridMonthlyInDate.AllowUserToAddRows = false
	  Me.gridMonthlyInDate.AllowUserToDeleteRows = false
	  Me.gridMonthlyInDate.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom)  _
	  	  	  Or System.Windows.Forms.AnchorStyles.Left)  _
	  	  	  Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
	  Me.gridMonthlyInDate.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
	  Me.gridMonthlyInDate.Location = New System.Drawing.Point(3, 31)
	  Me.gridMonthlyInDate.Name = "gridMonthlyInDate"
	  Me.gridMonthlyInDate.ReadOnly = true
	  Me.gridMonthlyInDate.RowTemplate.Height = 21
	  Me.gridMonthlyInDate.Size = New System.Drawing.Size(735, 234)
	  Me.gridMonthlyInDate.TabIndex = 4
	  '
	  'cboxMonthly
	  '
	  Me.cboxMonthly.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
	  Me.cboxMonthly.FormattingEnabled = true
	  Me.cboxMonthly.Location = New System.Drawing.Point(560, 6)
	  Me.cboxMonthly.Name = "cboxMonthly"
	  Me.cboxMonthly.Size = New System.Drawing.Size(160, 20)
	  Me.cboxMonthly.TabIndex = 1
	  Me.cboxMonthly.Text = "参照する月を選んでください"
	  AddHandler Me.cboxMonthly.SelectedIndexChanged, AddressOf Me.CboxMonthlySelectedIndexChanged
	  '
	  'tPageYearlyInDate
	  '
	  Me.tPageYearlyInDate.BackColor = System.Drawing.SystemColors.Control
	  Me.tPageYearlyInDate.Controls.Add(Me.gridTallyInDate)
	  Me.tPageYearlyInDate.Location = New System.Drawing.Point(4, 22)
	  Me.tPageYearlyInDate.Name = "tPageYearlyInDate"
	  Me.tPageYearlyInDate.Padding = New System.Windows.Forms.Padding(3)
	  Me.tPageYearlyInDate.Size = New System.Drawing.Size(741, 268)
	  Me.tPageYearlyInDate.TabIndex = 3
	  Me.tPageYearlyInDate.Text = "合計"
	  '
	  'gridTallyInDate
	  '
	  Me.gridTallyInDate.AllowUserToAddRows = false
	  Me.gridTallyInDate.AllowUserToDeleteRows = false
	  Me.gridTallyInDate.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom)  _
	  	  	  Or System.Windows.Forms.AnchorStyles.Left)  _
	  	  	  Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
	  Me.gridTallyInDate.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
	  Me.gridTallyInDate.Location = New System.Drawing.Point(3, 31)
	  Me.gridTallyInDate.Name = "gridTallyInDate"
	  Me.gridTallyInDate.ReadOnly = true
	  Me.gridTallyInDate.RowTemplate.Height = 21
	  Me.gridTallyInDate.Size = New System.Drawing.Size(735, 234)
	  Me.gridTallyInDate.TabIndex = 4
	  '
	  'tPageTally
	  '
	  Me.tPageTally.BackColor = System.Drawing.SystemColors.Control
	  Me.tPageTally.Controls.Add(Me.tabInTallyTab)
	  Me.tPageTally.Location = New System.Drawing.Point(4, 22)
	  Me.tPageTally.Name = "tPageTally"
	  Me.tPageTally.Padding = New System.Windows.Forms.Padding(3)
	  Me.tPageTally.Size = New System.Drawing.Size(763, 329)
	  Me.tPageTally.TabIndex = 2
	  Me.tPageTally.Text = "集計データ"
	  '
	  'tabInTallyTab
	  '
	  Me.tabInTallyTab.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom)  _
	  	  	  Or System.Windows.Forms.AnchorStyles.Left)  _
	  	  	  Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
	  Me.tabInTallyTab.Controls.Add(Me.tPageDailyInTally)
	  Me.tabInTallyTab.Controls.Add(Me.tPageWeeklyInTally)
	  Me.tabInTallyTab.Controls.Add(Me.tPageMonthlyInTally)
	  Me.tabInTallyTab.Location = New System.Drawing.Point(8, 29)
	  Me.tabInTallyTab.Name = "tabInTallyTab"
	  Me.tabInTallyTab.SelectedIndex = 0
	  Me.tabInTallyTab.Size = New System.Drawing.Size(749, 294)
	  Me.tabInTallyTab.TabIndex = 3
	  AddHandler Me.tabInTallyTab.SelectedIndexChanged, AddressOf Me.TabInTallyTab_SelectedIndexChanged
	  '
	  'tPageDailyInTally
	  '
	  Me.tPageDailyInTally.BackColor = System.Drawing.SystemColors.Control
	  Me.tPageDailyInTally.Controls.Add(Me.dataGridView1)
	  Me.tPageDailyInTally.Controls.Add(Me.cboxTallyMonthly)
	  Me.tPageDailyInTally.Location = New System.Drawing.Point(4, 22)
	  Me.tPageDailyInTally.Name = "tPageDailyInTally"
	  Me.tPageDailyInTally.Padding = New System.Windows.Forms.Padding(3)
	  Me.tPageDailyInTally.Size = New System.Drawing.Size(741, 268)
	  Me.tPageDailyInTally.TabIndex = 0
	  Me.tPageDailyInTally.Text = "日"
	  '
	  'dataGridView1
	  '
	  Me.dataGridView1.AllowUserToAddRows = false
	  Me.dataGridView1.AllowUserToDeleteRows = false
	  Me.dataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom)  _
	  	  	  Or System.Windows.Forms.AnchorStyles.Left)  _
	  	  	  Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
	  Me.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
	  Me.dataGridView1.Location = New System.Drawing.Point(3, 31)
	  Me.dataGridView1.Name = "dataGridView1"
	  Me.dataGridView1.ReadOnly = true
	  Me.dataGridView1.RowTemplate.Height = 21
	  Me.dataGridView1.Size = New System.Drawing.Size(735, 234)
	  Me.dataGridView1.TabIndex = 4
	  '
	  'cboxTallyMonthly
	  '
	  Me.cboxTallyMonthly.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
	  Me.cboxTallyMonthly.FormattingEnabled = true
	  Me.cboxTallyMonthly.Location = New System.Drawing.Point(560, 6)
	  Me.cboxTallyMonthly.Name = "cboxTallyMonthly"
	  Me.cboxTallyMonthly.Size = New System.Drawing.Size(160, 20)
	  Me.cboxTallyMonthly.TabIndex = 0
	  Me.cboxTallyMonthly.Text = "参照する月を選んでください"
	  AddHandler Me.cboxTallyMonthly.SelectedIndexChanged, AddressOf Me.CboxTallyMonthlySelectedIndexChanged
	  '
	  'tPageWeeklyInTally
	  '
	  Me.tPageWeeklyInTally.BackColor = System.Drawing.SystemColors.Control
	  Me.tPageWeeklyInTally.Controls.Add(Me.dataGridView2)
	  Me.tPageWeeklyInTally.Location = New System.Drawing.Point(4, 22)
	  Me.tPageWeeklyInTally.Name = "tPageWeeklyInTally"
	  Me.tPageWeeklyInTally.Padding = New System.Windows.Forms.Padding(3)
	  Me.tPageWeeklyInTally.Size = New System.Drawing.Size(741, 268)
	  Me.tPageWeeklyInTally.TabIndex = 1
	  Me.tPageWeeklyInTally.Text = "週"
	  '
	  'dataGridView2
	  '
	  Me.dataGridView2.AllowUserToAddRows = false
	  Me.dataGridView2.AllowUserToDeleteRows = false
	  Me.dataGridView2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom)  _
	  	  	  Or System.Windows.Forms.AnchorStyles.Left)  _
	  	  	  Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
	  Me.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
	  Me.dataGridView2.Location = New System.Drawing.Point(3, 31)
	  Me.dataGridView2.Name = "dataGridView2"
	  Me.dataGridView2.ReadOnly = true
	  Me.dataGridView2.RowTemplate.Height = 21
	  Me.dataGridView2.Size = New System.Drawing.Size(735, 234)
	  Me.dataGridView2.TabIndex = 5
	  '
	  'tPageMonthlyInTally
	  '
	  Me.tPageMonthlyInTally.BackColor = System.Drawing.SystemColors.Control
	  Me.tPageMonthlyInTally.Controls.Add(Me.dataGridView3)
	  Me.tPageMonthlyInTally.Location = New System.Drawing.Point(4, 22)
	  Me.tPageMonthlyInTally.Name = "tPageMonthlyInTally"
	  Me.tPageMonthlyInTally.Padding = New System.Windows.Forms.Padding(3)
	  Me.tPageMonthlyInTally.Size = New System.Drawing.Size(741, 268)
	  Me.tPageMonthlyInTally.TabIndex = 2
	  Me.tPageMonthlyInTally.Text = "月"
	  '
	  'dataGridView3
	  '
	  Me.dataGridView3.AllowUserToAddRows = false
	  Me.dataGridView3.AllowUserToDeleteRows = false
	  Me.dataGridView3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom)  _
	  	  	  Or System.Windows.Forms.AnchorStyles.Left)  _
	  	  	  Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
	  Me.dataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
	  Me.dataGridView3.Location = New System.Drawing.Point(3, 31)
	  Me.dataGridView3.Name = "dataGridView3"
	  Me.dataGridView3.ReadOnly = true
	  Me.dataGridView3.RowTemplate.Height = 21
	  Me.dataGridView3.Size = New System.Drawing.Size(735, 234)
	  Me.dataGridView3.TabIndex = 5
	  '
	  'btnClose
	  '
	  Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
	  Me.btnClose.Location = New System.Drawing.Point(651, 358)
	  Me.btnClose.Name = "btnClose"
	  Me.btnClose.Size = New System.Drawing.Size(110, 23)
	  Me.btnClose.TabIndex = 1
	  Me.btnClose.Text = "閉じる"
	  Me.btnClose.UseVisualStyleBackColor = true
	  AddHandler Me.btnClose.Click, AddressOf Me.BtnCloseClick
	  '
	  'btnReload
	  '
	  Me.btnReload.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
	  Me.btnReload.Location = New System.Drawing.Point(535, 358)
	  Me.btnReload.Name = "btnReload"
	  Me.btnReload.Size = New System.Drawing.Size(110, 23)
	  Me.btnReload.TabIndex = 2
	  Me.btnReload.Text = "全ファイル読み込み"
	  Me.btnReload.UseVisualStyleBackColor = true
	  '
	  'btnOutputCsv
	  '
	  Me.btnOutputCsv.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
	  Me.btnOutputCsv.Location = New System.Drawing.Point(419, 358)
	  Me.btnOutputCsv.Name = "btnOutputCsv"
	  Me.btnOutputCsv.Size = New System.Drawing.Size(110, 23)
	  Me.btnOutputCsv.TabIndex = 3
	  Me.btnOutputCsv.Text = "CSV出力"
	  Me.btnOutputCsv.UseVisualStyleBackColor = true
	  AddHandler Me.btnOutputCsv.Click, AddressOf Me.BtnOutputCsvClick
	  '
	  'chkBoxExcludeData
	  '
	  Me.chkBoxExcludeData.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left),System.Windows.Forms.AnchorStyles)
	  Me.chkBoxExcludeData.AutoSize = true
	  Me.chkBoxExcludeData.Checked = true
	  Me.chkBoxExcludeData.CheckState = System.Windows.Forms.CheckState.Checked
	  Me.chkBoxExcludeData.Location = New System.Drawing.Point(12, 362)
	  Me.chkBoxExcludeData.Name = "chkBoxExcludeData"
	  Me.chkBoxExcludeData.Size = New System.Drawing.Size(234, 16)
	  Me.chkBoxExcludeData.TabIndex = 4
	  Me.chkBoxExcludeData.Text = "合計件数から作業時間記入漏れの日を除く"
	  Me.chkBoxExcludeData.UseVisualStyleBackColor = true
	  AddHandler Me.chkBoxExcludeData.CheckedChanged, AddressOf Me.ChkBoxExcludeDataCheckedChanged
	  '
	  'MainForm
	  '
	  Me.AutoScaleDimensions = New System.Drawing.SizeF(6!, 12!)
	  Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
	  Me.ClientSize = New System.Drawing.Size(769, 393)
	  Me.Controls.Add(Me.chkBoxExcludeData)
	  Me.Controls.Add(Me.btnOutputCsv)
	  Me.Controls.Add(Me.btnReload)
	  Me.Controls.Add(Me.btnClose)
	  Me.Controls.Add(Me.tabRoot)
	  Me.Name = "MainForm"
	  Me.Text = "WorkReportAnalysis"
	  AddHandler Closing, AddressOf Me.MainFormClosing
	  AddHandler Load, AddressOf Me.MainFormLoad
	  Me.tabRoot.ResumeLayout(false)
	  Me.tPagePersonal.ResumeLayout(false)
	  Me.tabInPersonalTab.ResumeLayout(false)
	  Me.tPageMonth10InPersonal.ResumeLayout(false)
	  CType(Me.gridMonth10InPersonal,System.ComponentModel.ISupportInitialize).EndInit
	  Me.tPageDate.ResumeLayout(false)
	  Me.tabInDateTab.ResumeLayout(false)
	  Me.tPageDailyInDate.ResumeLayout(false)
	  CType(Me.gridDailyInDate,System.ComponentModel.ISupportInitialize).EndInit
	  Me.tPageWeeklyInDate.ResumeLayout(false)
	  CType(Me.gridWeeklyInDate,System.ComponentModel.ISupportInitialize).EndInit
	  Me.tPageMonthlyInDate.ResumeLayout(false)
	  CType(Me.gridMonthlyInDate,System.ComponentModel.ISupportInitialize).EndInit
	  Me.tPageYearlyInDate.ResumeLayout(false)
	  CType(Me.gridTallyInDate,System.ComponentModel.ISupportInitialize).EndInit
	  Me.tPageTally.ResumeLayout(false)
	  Me.tabInTallyTab.ResumeLayout(false)
	  Me.tPageDailyInTally.ResumeLayout(false)
	  CType(Me.dataGridView1,System.ComponentModel.ISupportInitialize).EndInit
	  Me.tPageWeeklyInTally.ResumeLayout(false)
	  CType(Me.dataGridView2,System.ComponentModel.ISupportInitialize).EndInit
	  Me.tPageMonthlyInTally.ResumeLayout(false)
	  CType(Me.dataGridView3,System.ComponentModel.ISupportInitialize).EndInit
	  Me.ResumeLayout(false)
	  Me.PerformLayout
	End Sub
	Private dataGridView3 As System.Windows.Forms.DataGridView
	Private dataGridView2 As System.Windows.Forms.DataGridView
	Private dataGridView1 As System.Windows.Forms.DataGridView
	Private gridTallyInDate As System.Windows.Forms.DataGridView
	Private gridMonthlyInDate As System.Windows.Forms.DataGridView
	Private gridWeeklyInDate As System.Windows.Forms.DataGridView
	Private gridDailyInDate As System.Windows.Forms.DataGridView
	Private dateTimePickerInDatePage As System.Windows.Forms.DateTimePicker
	Private gridMonth10InPersonal As System.Windows.Forms.DataGridView
	Private cboxMonthly As System.Windows.Forms.ComboBox
	Private cboxWeekly As System.Windows.Forms.ComboBox
	Private cboxTallyMonthly As System.Windows.Forms.ComboBox
	Private tPageMonthlyInTally As System.Windows.Forms.TabPage
	Private tPageWeeklyInTally As System.Windows.Forms.TabPage
	Private tPageDailyInTally As System.Windows.Forms.TabPage
	Private tabInTallyTab As System.Windows.Forms.TabControl
	Private tPageYearlyInDate As System.Windows.Forms.TabPage
	Private tPageMonthlyInDate As System.Windows.Forms.TabPage
	Private tPageWeeklyInDate As System.Windows.Forms.TabPage
	Private tPageDailyInDate As System.Windows.Forms.TabPage
	Private tabInDateTab As System.Windows.Forms.TabControl
	Private chkBoxExcludeData As System.Windows.Forms.CheckBox
	Private btnOutputCsv As System.Windows.Forms.Button
	Private btnReload As System.Windows.Forms.Button
	Private cboxUserName As System.Windows.Forms.ComboBox
	Private tPageMonth10InPersonal As System.Windows.Forms.TabPage
	Private tabInPersonalTab As System.Windows.Forms.TabControl
	Private tPageTally As System.Windows.Forms.TabPage
	Private btnClose As System.Windows.Forms.Button
	Private tPageDate As System.Windows.Forms.TabPage
	Private tPagePersonal As System.Windows.Forms.TabPage
	Private tabRoot As System.Windows.Forms.TabControl
End Class
