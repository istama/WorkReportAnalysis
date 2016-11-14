'
' 日付: 2016/06/20
'
Partial Class ProgressBarForm
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
		Me.pBar = New System.Windows.Forms.ProgressBar()
		Me.btnCancel = New System.Windows.Forms.Button()
		Me.lblMsg = New System.Windows.Forms.Label()
		Me.backgroundWorker1 = New System.ComponentModel.BackgroundWorker()
		Me.SuspendLayout
		'
		'pBar
		'
		Me.pBar.Location = New System.Drawing.Point(13, 29)
		Me.pBar.Name = "pBar"
		Me.pBar.Size = New System.Drawing.Size(259, 23)
		Me.pBar.TabIndex = 0
		'
		'btnCancel
		'
		Me.btnCancel.Location = New System.Drawing.Point(197, 58)
		Me.btnCancel.Name = "btnCancel"
		Me.btnCancel.Size = New System.Drawing.Size(75, 23)
		Me.btnCancel.TabIndex = 1
		Me.btnCancel.Text = "キャンセル"
		Me.btnCancel.UseVisualStyleBackColor = true
		AddHandler Me.btnCancel.Click, AddressOf Me.BtnCancelClick
		'
		'lblMsg
		'
		Me.lblMsg.AutoSize = true
		Me.lblMsg.Location = New System.Drawing.Point(13, 9)
		Me.lblMsg.Name = "lblMsg"
		Me.lblMsg.Size = New System.Drawing.Size(35, 12)
		Me.lblMsg.TabIndex = 2
		Me.lblMsg.Text = "label1"
		'
		'backgroundWorker1
		'
		AddHandler Me.backgroundWorker1.DoWork, AddressOf Me.BackgroundWorker1DoWork
		AddHandler Me.backgroundWorker1.RunWorkerCompleted, AddressOf Me.BackgroundWorker1RunWorkerCompleted
		'
		'ProgressBarForm
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6!, 12!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(284, 90)
		Me.Controls.Add(Me.lblMsg)
		Me.Controls.Add(Me.btnCancel)
		Me.Controls.Add(Me.pBar)
		Me.MaximizeBox = false
		Me.Name = "ProgressBarForm"
		Me.Text = "読み込み中..."
		AddHandler Load, AddressOf Me.ProgressBarFormLoad
		Me.ResumeLayout(false)
		Me.PerformLayout
	End Sub
	Private backgroundWorker1 As System.ComponentModel.BackgroundWorker
	Private lblMsg As System.Windows.Forms.Label
	Private btnCancel As System.Windows.Forms.Button
	Private pBar As System.Windows.Forms.ProgressBar
	

End Class
