'
' 日付: 2016/05/09
'
Public Class MsgBox	
	
	Public Shared Sub Show(text As String, Optional title As String="")
		MessageBox.Show(text, title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
	End Sub
	
	Public Shared Sub ShowError(ex As Exception)
		MessageBox.Show(ex.ToString, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning)
	End Sub
End Class
