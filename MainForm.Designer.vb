<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MainForm
  Inherits System.Windows.Forms.Form

  'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
  <System.Diagnostics.DebuggerNonUserCode()>
  Protected Overrides Sub Dispose(ByVal disposing As Boolean)
    Try
      If disposing AndAlso components IsNot Nothing Then
        components.Dispose()
      End If
    Finally
      MyBase.Dispose(disposing)
    End Try
  End Sub

  'Windows フォーム デザイナーで必要です。
  Private components As System.ComponentModel.IContainer

  'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
  'Windows フォーム デザイナーを使用して変更できます。  
  'コード エディターを使って変更しないでください。
  <System.Diagnostics.DebuggerStepThrough()>
  Private Sub InitializeComponent()
    Me.TabPage2 = New System.Windows.Forms.TabPage()
    Me.tabInTermTab = New System.Windows.Forms.TabControl()
    Me.pageDays = New System.Windows.Forms.TabPage()
    Me.dPicDailyTerm = New System.Windows.Forms.DateTimePicker()
    Me.pageWeeks = New System.Windows.Forms.TabPage()
    Me.cboxWeeklyTerm = New System.Windows.Forms.ComboBox()
    Me.pageMonths = New System.Windows.Forms.TabPage()
    Me.cboxMonthlyTerm = New System.Windows.Forms.ComboBox()
    Me.pageYear = New System.Windows.Forms.TabPage()
    Me.TabPage1 = New System.Windows.Forms.TabPage()
    Me.cboxUserInfo = New System.Windows.Forms.ComboBox()
    Me.tabInPersonalTab = New System.Windows.Forms.TabControl()
    Me.page10Month = New System.Windows.Forms.TabPage()
    Me.page11Month = New System.Windows.Forms.TabPage()
    Me.page12Month = New System.Windows.Forms.TabPage()
    Me.pageSum = New System.Windows.Forms.TabPage()
    Me.btnClose = New System.Windows.Forms.Button()
    Me.tabMaster = New System.Windows.Forms.TabControl()
    Me.TabPage3 = New System.Windows.Forms.TabPage()
    Me.tabInTotalTab = New System.Windows.Forms.TabControl()
    Me.pageDailyTotal = New System.Windows.Forms.TabPage()
    Me.cboxDailyTotal = New System.Windows.Forms.ComboBox()
    Me.pageWeeklyTotal = New System.Windows.Forms.TabPage()
    Me.pageMonthlyTotal = New System.Windows.Forms.TabPage()
    Me.cmdReadAllFile = New System.Windows.Forms.Button()
    Me.chkExcludeIncompleteRecordFromSum = New System.Windows.Forms.CheckBox()
    Me.cmdOutputCSV = New System.Windows.Forms.Button()
    Me.TabPage2.SuspendLayout()
    Me.tabInTermTab.SuspendLayout()
    Me.pageDays.SuspendLayout()
    Me.pageWeeks.SuspendLayout()
    Me.pageMonths.SuspendLayout()
    Me.TabPage1.SuspendLayout()
    Me.tabInPersonalTab.SuspendLayout()
    Me.tabMaster.SuspendLayout()
    Me.TabPage3.SuspendLayout()
    Me.tabInTotalTab.SuspendLayout()
    Me.pageDailyTotal.SuspendLayout()
    Me.SuspendLayout()
    '
    'TabPage2
    '
    Me.TabPage2.AutoScroll = True
    Me.TabPage2.BackColor = System.Drawing.SystemColors.Control
    Me.TabPage2.Controls.Add(Me.tabInTermTab)
    Me.TabPage2.Location = New System.Drawing.Point(4, 22)
    Me.TabPage2.Name = "TabPage2"
    Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
    Me.TabPage2.Size = New System.Drawing.Size(1274, 707)
    Me.TabPage2.TabIndex = 1
    Me.TabPage2.Text = "日付データ"
    '
    'tabInTermTab
    '
    Me.tabInTermTab.Controls.Add(Me.pageDays)
    Me.tabInTermTab.Controls.Add(Me.pageWeeks)
    Me.tabInTermTab.Controls.Add(Me.pageMonths)
    Me.tabInTermTab.Controls.Add(Me.pageYear)
    Me.tabInTermTab.Location = New System.Drawing.Point(18, 32)
    Me.tabInTermTab.Name = "tabInTermTab"
    Me.tabInTermTab.SelectedIndex = 0
    Me.tabInTermTab.Size = New System.Drawing.Size(1235, 652)
    Me.tabInTermTab.TabIndex = 0
    '
    'pageDays
    '
    Me.pageDays.BackColor = System.Drawing.SystemColors.Control
    Me.pageDays.Controls.Add(Me.dPicDailyTerm)
    Me.pageDays.Location = New System.Drawing.Point(4, 22)
    Me.pageDays.Name = "pageDays"
    Me.pageDays.Padding = New System.Windows.Forms.Padding(3)
    Me.pageDays.Size = New System.Drawing.Size(1227, 626)
    Me.pageDays.TabIndex = 0
    Me.pageDays.Text = "日"
    '
    'dPicDailyTerm
    '
    Me.dPicDailyTerm.Location = New System.Drawing.Point(950, 15)
    Me.dPicDailyTerm.Name = "dPicDailyTerm"
    Me.dPicDailyTerm.Size = New System.Drawing.Size(244, 19)
    Me.dPicDailyTerm.TabIndex = 0
    '
    'pageWeeks
    '
    Me.pageWeeks.BackColor = System.Drawing.SystemColors.Control
    Me.pageWeeks.Controls.Add(Me.cboxWeeklyTerm)
    Me.pageWeeks.Location = New System.Drawing.Point(4, 22)
    Me.pageWeeks.Name = "pageWeeks"
    Me.pageWeeks.Padding = New System.Windows.Forms.Padding(3)
    Me.pageWeeks.Size = New System.Drawing.Size(1227, 626)
    Me.pageWeeks.TabIndex = 1
    Me.pageWeeks.Text = "週"
    '
    'cboxWeeklyTerm
    '
    Me.cboxWeeklyTerm.FormattingEnabled = True
    Me.cboxWeeklyTerm.Location = New System.Drawing.Point(950, 15)
    Me.cboxWeeklyTerm.Name = "cboxWeeklyTerm"
    Me.cboxWeeklyTerm.Size = New System.Drawing.Size(244, 20)
    Me.cboxWeeklyTerm.TabIndex = 0
    Me.cboxWeeklyTerm.Text = "参照する週を選んで下さい"
    '
    'pageMonths
    '
    Me.pageMonths.BackColor = System.Drawing.SystemColors.Control
    Me.pageMonths.Controls.Add(Me.cboxMonthlyTerm)
    Me.pageMonths.Location = New System.Drawing.Point(4, 22)
    Me.pageMonths.Name = "pageMonths"
    Me.pageMonths.Padding = New System.Windows.Forms.Padding(3)
    Me.pageMonths.Size = New System.Drawing.Size(1227, 626)
    Me.pageMonths.TabIndex = 2
    Me.pageMonths.Text = "月"
    '
    'cboxMonthlyTerm
    '
    Me.cboxMonthlyTerm.FormattingEnabled = True
    Me.cboxMonthlyTerm.Location = New System.Drawing.Point(950, 15)
    Me.cboxMonthlyTerm.Name = "cboxMonthlyTerm"
    Me.cboxMonthlyTerm.Size = New System.Drawing.Size(244, 20)
    Me.cboxMonthlyTerm.TabIndex = 1
    Me.cboxMonthlyTerm.Text = "参照する月を選んで下さい"
    '
    'pageYear
    '
    Me.pageYear.BackColor = System.Drawing.SystemColors.Control
    Me.pageYear.Location = New System.Drawing.Point(4, 22)
    Me.pageYear.Name = "pageYear"
    Me.pageYear.Padding = New System.Windows.Forms.Padding(3)
    Me.pageYear.Size = New System.Drawing.Size(1227, 626)
    Me.pageYear.TabIndex = 3
    Me.pageYear.Text = "合計"
    '
    'TabPage1
    '
    Me.TabPage1.AutoScroll = True
    Me.TabPage1.BackColor = System.Drawing.SystemColors.Control
    Me.TabPage1.Controls.Add(Me.cboxUserInfo)
    Me.TabPage1.Controls.Add(Me.tabInPersonalTab)
    Me.TabPage1.Location = New System.Drawing.Point(4, 22)
    Me.TabPage1.Name = "TabPage1"
    Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
    Me.TabPage1.Size = New System.Drawing.Size(1274, 707)
    Me.TabPage1.TabIndex = 0
    Me.TabPage1.Text = "個人データ"
    '
    'cboxUserInfo
    '
    Me.cboxUserInfo.FormattingEnabled = True
    Me.cboxUserInfo.Location = New System.Drawing.Point(946, 6)
    Me.cboxUserInfo.Name = "cboxUserInfo"
    Me.cboxUserInfo.Size = New System.Drawing.Size(271, 20)
    Me.cboxUserInfo.TabIndex = 1
    Me.cboxUserInfo.Text = "参照ユーザを選んでください"
    '
    'tabInPersonalTab
    '
    Me.tabInPersonalTab.Controls.Add(Me.page10Month)
    Me.tabInPersonalTab.Controls.Add(Me.page11Month)
    Me.tabInPersonalTab.Controls.Add(Me.page12Month)
    Me.tabInPersonalTab.Controls.Add(Me.pageSum)
    Me.tabInPersonalTab.Location = New System.Drawing.Point(18, 32)
    Me.tabInPersonalTab.Name = "tabInPersonalTab"
    Me.tabInPersonalTab.SelectedIndex = 0
    Me.tabInPersonalTab.Size = New System.Drawing.Size(1235, 652)
    Me.tabInPersonalTab.TabIndex = 0
    '
    'page10Month
    '
    Me.page10Month.BackColor = System.Drawing.SystemColors.Control
    Me.page10Month.Location = New System.Drawing.Point(4, 22)
    Me.page10Month.Name = "page10Month"
    Me.page10Month.Padding = New System.Windows.Forms.Padding(3)
    Me.page10Month.Size = New System.Drawing.Size(1227, 626)
    Me.page10Month.TabIndex = 0
    Me.page10Month.Text = "10月"
    '
    'page11Month
    '
    Me.page11Month.BackColor = System.Drawing.SystemColors.Control
    Me.page11Month.Location = New System.Drawing.Point(4, 22)
    Me.page11Month.Name = "page11Month"
    Me.page11Month.Padding = New System.Windows.Forms.Padding(3)
    Me.page11Month.Size = New System.Drawing.Size(1227, 626)
    Me.page11Month.TabIndex = 1
    Me.page11Month.Text = "11月"
    '
    'page12Month
    '
    Me.page12Month.BackColor = System.Drawing.SystemColors.Control
    Me.page12Month.Location = New System.Drawing.Point(4, 22)
    Me.page12Month.Name = "page12Month"
    Me.page12Month.Size = New System.Drawing.Size(1227, 626)
    Me.page12Month.TabIndex = 2
    Me.page12Month.Text = "12月"
    '
    'pageSum
    '
    Me.pageSum.BackColor = System.Drawing.SystemColors.Control
    Me.pageSum.Location = New System.Drawing.Point(4, 22)
    Me.pageSum.Name = "pageSum"
    Me.pageSum.Size = New System.Drawing.Size(1227, 626)
    Me.pageSum.TabIndex = 3
    Me.pageSum.Text = "集計"
    '
    'btnClose
    '
    Me.btnClose.Location = New System.Drawing.Point(1128, 744)
    Me.btnClose.Name = "btnClose"
    Me.btnClose.Size = New System.Drawing.Size(141, 32)
    Me.btnClose.TabIndex = 2
    Me.btnClose.Text = "閉じる"
    Me.btnClose.UseVisualStyleBackColor = True
    '
    'tabMaster
    '
    Me.tabMaster.Controls.Add(Me.TabPage1)
    Me.tabMaster.Controls.Add(Me.TabPage2)
    Me.tabMaster.Controls.Add(Me.TabPage3)
    Me.tabMaster.Location = New System.Drawing.Point(0, 0)
    Me.tabMaster.Name = "tabMaster"
    Me.tabMaster.SelectedIndex = 0
    Me.tabMaster.Size = New System.Drawing.Size(1282, 733)
    Me.tabMaster.TabIndex = 0
    '
    'TabPage3
    '
    Me.TabPage3.AutoScroll = True
    Me.TabPage3.Controls.Add(Me.tabInTotalTab)
    Me.TabPage3.Location = New System.Drawing.Point(4, 22)
    Me.TabPage3.Name = "TabPage3"
    Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
    Me.TabPage3.Size = New System.Drawing.Size(1274, 707)
    Me.TabPage3.TabIndex = 2
    Me.TabPage3.Text = "集計データ"
    '
    'tabInTotalTab
    '
    Me.tabInTotalTab.Controls.Add(Me.pageDailyTotal)
    Me.tabInTotalTab.Controls.Add(Me.pageWeeklyTotal)
    Me.tabInTotalTab.Controls.Add(Me.pageMonthlyTotal)
    Me.tabInTotalTab.Location = New System.Drawing.Point(18, 32)
    Me.tabInTotalTab.Name = "tabInTotalTab"
    Me.tabInTotalTab.SelectedIndex = 0
    Me.tabInTotalTab.Size = New System.Drawing.Size(1235, 652)
    Me.tabInTotalTab.TabIndex = 1
    '
    'pageDailyTotal
    '
    Me.pageDailyTotal.BackColor = System.Drawing.SystemColors.Control
    Me.pageDailyTotal.Controls.Add(Me.cboxDailyTotal)
    Me.pageDailyTotal.Location = New System.Drawing.Point(4, 22)
    Me.pageDailyTotal.Name = "pageDailyTotal"
    Me.pageDailyTotal.Padding = New System.Windows.Forms.Padding(3)
    Me.pageDailyTotal.Size = New System.Drawing.Size(1227, 626)
    Me.pageDailyTotal.TabIndex = 0
    Me.pageDailyTotal.Text = "日"
    '
    'cboxDailyTotal
    '
    Me.cboxDailyTotal.FormattingEnabled = True
    Me.cboxDailyTotal.Location = New System.Drawing.Point(950, 15)
    Me.cboxDailyTotal.Name = "cboxDailyTotal"
    Me.cboxDailyTotal.Size = New System.Drawing.Size(244, 20)
    Me.cboxDailyTotal.TabIndex = 2
    Me.cboxDailyTotal.Text = "参照する月を選んで下さい"
    '
    'pageWeeklyTotal
    '
    Me.pageWeeklyTotal.BackColor = System.Drawing.SystemColors.Control
    Me.pageWeeklyTotal.Location = New System.Drawing.Point(4, 22)
    Me.pageWeeklyTotal.Name = "pageWeeklyTotal"
    Me.pageWeeklyTotal.Padding = New System.Windows.Forms.Padding(3)
    Me.pageWeeklyTotal.Size = New System.Drawing.Size(1227, 626)
    Me.pageWeeklyTotal.TabIndex = 1
    Me.pageWeeklyTotal.Text = "週"
    '
    'pageMonthlyTotal
    '
    Me.pageMonthlyTotal.BackColor = System.Drawing.SystemColors.Control
    Me.pageMonthlyTotal.Location = New System.Drawing.Point(4, 22)
    Me.pageMonthlyTotal.Name = "pageMonthlyTotal"
    Me.pageMonthlyTotal.Padding = New System.Windows.Forms.Padding(3)
    Me.pageMonthlyTotal.Size = New System.Drawing.Size(1227, 626)
    Me.pageMonthlyTotal.TabIndex = 2
    Me.pageMonthlyTotal.Text = "月"
    '
    'cmdReadAllFile
    '
    Me.cmdReadAllFile.Location = New System.Drawing.Point(981, 744)
    Me.cmdReadAllFile.Name = "cmdReadAllFile"
    Me.cmdReadAllFile.Size = New System.Drawing.Size(141, 32)
    Me.cmdReadAllFile.TabIndex = 3
    Me.cmdReadAllFile.Text = "全ファイル読み込み"
    Me.cmdReadAllFile.UseVisualStyleBackColor = True
    '
    'chkExcludeIncompleteRecordFromSum
    '
    Me.chkExcludeIncompleteRecordFromSum.AutoSize = True
    Me.chkExcludeIncompleteRecordFromSum.Checked = True
    Me.chkExcludeIncompleteRecordFromSum.CheckState = System.Windows.Forms.CheckState.Checked
    Me.chkExcludeIncompleteRecordFromSum.Location = New System.Drawing.Point(22, 740)
    Me.chkExcludeIncompleteRecordFromSum.Name = "chkExcludeIncompleteRecordFromSum"
    Me.chkExcludeIncompleteRecordFromSum.Size = New System.Drawing.Size(234, 16)
    Me.chkExcludeIncompleteRecordFromSum.TabIndex = 4
    Me.chkExcludeIncompleteRecordFromSum.Text = "合計件数から作業時間記入漏れの日を除く"
    Me.chkExcludeIncompleteRecordFromSum.UseVisualStyleBackColor = True
    '
    'cmdOutputCSV
    '
    Me.cmdOutputCSV.Location = New System.Drawing.Point(834, 744)
    Me.cmdOutputCSV.Name = "cmdOutputCSV"
    Me.cmdOutputCSV.Size = New System.Drawing.Size(141, 32)
    Me.cmdOutputCSV.TabIndex = 5
    Me.cmdOutputCSV.Text = "CSV出力"
    Me.cmdOutputCSV.UseVisualStyleBackColor = True
    '
    'MainForm
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.BackColor = System.Drawing.SystemColors.Control
    Me.ClientSize = New System.Drawing.Size(1281, 788)
    Me.Controls.Add(Me.cmdOutputCSV)
    Me.Controls.Add(Me.chkExcludeIncompleteRecordFromSum)
    Me.Controls.Add(Me.cmdReadAllFile)
    Me.Controls.Add(Me.btnClose)
    Me.Controls.Add(Me.tabMaster)
    Me.Name = "MainForm"
    Me.Text = "WorkReportAnalysis"
    Me.TabPage2.ResumeLayout(False)
    Me.tabInTermTab.ResumeLayout(False)
    Me.pageDays.ResumeLayout(False)
    Me.pageWeeks.ResumeLayout(False)
    Me.pageMonths.ResumeLayout(False)
    Me.TabPage1.ResumeLayout(False)
    Me.tabInPersonalTab.ResumeLayout(False)
    Me.tabMaster.ResumeLayout(False)
    Me.TabPage3.ResumeLayout(False)
    Me.tabInTotalTab.ResumeLayout(False)
    Me.pageDailyTotal.ResumeLayout(False)
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents TabPage2 As TabPage
  Friend WithEvents TabPage1 As TabPage
  Friend WithEvents cboxUserInfo As ComboBox
  Friend WithEvents tabInPersonalTab As TabControl
  Friend WithEvents page10Month As TabPage
  Friend WithEvents page11Month As TabPage
  Friend WithEvents page12Month As TabPage
  Friend WithEvents pageSum As TabPage
  Friend WithEvents tabMaster As TabControl
  Friend WithEvents btnClose As Button
  Friend WithEvents tabInTermTab As TabControl
  Friend WithEvents pageDays As TabPage
  Friend WithEvents pageWeeks As TabPage
  Friend WithEvents pageMonths As TabPage
  Friend WithEvents pageYear As TabPage
  Friend WithEvents dPicDailyTerm As DateTimePicker
  Friend WithEvents cboxWeeklyTerm As ComboBox
  Friend WithEvents cmdReadAllFile As Button
  Friend WithEvents cboxMonthlyTerm As ComboBox
  Friend WithEvents chkExcludeIncompleteRecordFromSum As CheckBox
  Friend WithEvents TabPage3 As TabPage
  Friend WithEvents tabInTotalTab As TabControl
  Friend WithEvents pageDailyTotal As TabPage
  Friend WithEvents pageWeeklyTotal As TabPage
  Friend WithEvents pageMonthlyTotal As TabPage
  Friend WithEvents cboxDailyTotal As ComboBox
  Friend WithEvents cmdOutputCSV As Button
End Class
