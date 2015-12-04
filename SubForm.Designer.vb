<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class SubForm
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
    Me.chkExcludeIncompleteRecordFromSum = New System.Windows.Forms.CheckBox()
    Me.btnClose = New System.Windows.Forms.Button()
    Me.cmdOutputCSV = New System.Windows.Forms.Button()
    Me.SuspendLayout()
    '
    'chkExcludeIncompleteRecordFromSum
    '
    Me.chkExcludeIncompleteRecordFromSum.AutoSize = True
    Me.chkExcludeIncompleteRecordFromSum.Checked = True
    Me.chkExcludeIncompleteRecordFromSum.CheckState = System.Windows.Forms.CheckState.Checked
    Me.chkExcludeIncompleteRecordFromSum.Location = New System.Drawing.Point(12, 661)
    Me.chkExcludeIncompleteRecordFromSum.Name = "chkExcludeIncompleteRecordFromSum"
    Me.chkExcludeIncompleteRecordFromSum.Size = New System.Drawing.Size(234, 16)
    Me.chkExcludeIncompleteRecordFromSum.TabIndex = 2
    Me.chkExcludeIncompleteRecordFromSum.Text = "合計件数から作業時間記入漏れの日を除く"
    Me.chkExcludeIncompleteRecordFromSum.UseVisualStyleBackColor = True
    '
    'btnClose
    '
    Me.btnClose.Location = New System.Drawing.Point(1081, 661)
    Me.btnClose.Name = "btnClose"
    Me.btnClose.Size = New System.Drawing.Size(141, 32)
    Me.btnClose.TabIndex = 3
    Me.btnClose.Text = "閉じる"
    Me.btnClose.UseVisualStyleBackColor = True
    '
    'cmdOutputCSV
    '
    Me.cmdOutputCSV.Location = New System.Drawing.Point(934, 661)
    Me.cmdOutputCSV.Name = "cmdOutputCSV"
    Me.cmdOutputCSV.Size = New System.Drawing.Size(141, 32)
    Me.cmdOutputCSV.TabIndex = 6
    Me.cmdOutputCSV.Text = "CSV出力"
    Me.cmdOutputCSV.UseVisualStyleBackColor = True
    '
    'SubForm
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(1234, 704)
    Me.Controls.Add(Me.cmdOutputCSV)
    Me.Controls.Add(Me.btnClose)
    Me.Controls.Add(Me.chkExcludeIncompleteRecordFromSum)
    Me.Name = "SubForm"
    Me.Text = "PersonalForm"
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents chkExcludeIncompleteRecordFromSum As CheckBox
  Friend WithEvents btnClose As Button
  Friend WithEvents cmdOutputCSV As Button
End Class
