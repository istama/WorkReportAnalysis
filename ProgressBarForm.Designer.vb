<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ProgressBarForm
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
    Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
    Me.Label1 = New System.Windows.Forms.Label()
    Me.Button1 = New System.Windows.Forms.Button()
    Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
    Me.lblFileName = New System.Windows.Forms.Label()
    Me.SuspendLayout()
    '
    'ProgressBar1
    '
    Me.ProgressBar1.Location = New System.Drawing.Point(12, 28)
    Me.ProgressBar1.Name = "ProgressBar1"
    Me.ProgressBar1.Size = New System.Drawing.Size(260, 23)
    Me.ProgressBar1.TabIndex = 0
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(13, 13)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(132, 12)
    Me.Label1.TabIndex = 1
    Me.Label1.Text = "ファイルを読み込んでいます"
    '
    'Button1
    '
    Me.Button1.Location = New System.Drawing.Point(197, 57)
    Me.Button1.Name = "Button1"
    Me.Button1.Size = New System.Drawing.Size(75, 23)
    Me.Button1.TabIndex = 2
    Me.Button1.Text = "キャンセル"
    Me.Button1.UseVisualStyleBackColor = True
    '
    'BackgroundWorker1
    '
    '
    'lblFileName
    '
    Me.lblFileName.AutoSize = True
    Me.lblFileName.Location = New System.Drawing.Point(160, 13)
    Me.lblFileName.Name = "lblFileName"
    Me.lblFileName.Size = New System.Drawing.Size(0, 12)
    Me.lblFileName.TabIndex = 3
    '
    'ProgressBarForm
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(284, 91)
    Me.ControlBox = False
    Me.Controls.Add(Me.lblFileName)
    Me.Controls.Add(Me.Button1)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.ProgressBar1)
    Me.Name = "ProgressBarForm"
    Me.Text = "読み込み中..."
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub

  Friend WithEvents ProgressBar1 As ProgressBar
  Friend WithEvents Label1 As Label
  Friend WithEvents Button1 As Button
  Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
  Friend WithEvents lblFileName As Label
End Class
