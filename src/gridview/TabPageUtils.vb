'
' 日付: 2016/07/16
'
Public Class TabPageUtils
	Public Shared Function GetDataGridView(tabPage As TabPage) As DataGridView
		Dim grid As DataGridView = Nothing
		For i = 0 To tabPage.Controls.Count
			grid = TryCast(tabPage.Controls.Item(i), DataGridView)
			If grid IsNot Nothing Then
				Exit For
			End If
		Next
		
		Return grid
	End Function
	
	Public Shared Function GetComboBox(tabPage As TabPage) As ComboBox
		Dim cbox As ComboBox = Nothing
		For i = 0 To tabPage.Controls.Count
			cbox = TryCast(tabPage.Controls.Item(i), ComboBox)
			If cbox IsNot Nothing Then
				Exit For
			End If
		Next
		
		Return cbox		
	End Function
	
	Public Shared Function GetDateTimePicker(tabPage As TabPage) As DateTimePicker
		Dim dPicker As DateTimePicker = Nothing
		For i = 0 To tabPage.Controls.Count
			dPicker = TryCast(tabPage.Controls.Item(i), DateTimePicker)
			If dPicker IsNot Nothing Then
				Exit For
			End If
		Next
		
		Return dPicker		
	End Function	
End Class
