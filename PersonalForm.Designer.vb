<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PersonalForm
  Inherits System.Windows.Forms.Form

  'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
  <System.Diagnostics.DebuggerNonUserCode()> _
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
  <System.Diagnostics.DebuggerStepThrough()> _
  Private Sub InitializeComponent()
    Me.tabInPersonalTab = New System.Windows.Forms.TabControl()
    Me.page10Month = New System.Windows.Forms.TabPage()
    Me.page11Month = New System.Windows.Forms.TabPage()
    Me.page12Month = New System.Windows.Forms.TabPage()
    Me.pageSum = New System.Windows.Forms.TabPage()
    Me.tabInPersonalTab.SuspendLayout()
    Me.SuspendLayout()
    '
    'tabInPersonalTab
    '
    Me.tabInPersonalTab.Controls.Add(Me.page10Month)
    Me.tabInPersonalTab.Controls.Add(Me.page11Month)
    Me.tabInPersonalTab.Controls.Add(Me.page12Month)
    Me.tabInPersonalTab.Controls.Add(Me.pageSum)
    Me.tabInPersonalTab.Location = New System.Drawing.Point(0, 0)
    Me.tabInPersonalTab.Name = "tabInPersonalTab"
    Me.tabInPersonalTab.SelectedIndex = 0
    Me.tabInPersonalTab.Size = New System.Drawing.Size(1235, 652)
    Me.tabInPersonalTab.TabIndex = 1
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
    'PersonalForm
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(1234, 652)
    Me.Controls.Add(Me.tabInPersonalTab)
    Me.Name = "PersonalForm"
    Me.Text = "PersonalForm"
    Me.tabInPersonalTab.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub

  Friend WithEvents tabInPersonalTab As TabControl
  Friend WithEvents page10Month As TabPage
  Friend WithEvents page11Month As TabPage
  Friend WithEvents page12Month As TabPage
  Friend WithEvents pageSum As TabPage
End Class
