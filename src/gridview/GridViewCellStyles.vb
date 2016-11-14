'
' 日付: 2016/07/16
'
Imports Common.Util

''' <summary>
''' 共通して使用されるグリッドビューの表示スタイルをまとめた構造体。
''' </summary>
Public Structure MyGridViewCellStyles
	Private _holidayRowCellStyle As DataGridViewCellStyle
	Public ReadOnly Property HolidayRow As DataGridViewCellStyle
		Get
			Return _holidayRowCellStyle
		End Get
	End Property
	
	Private _monthlyTotalRowCellStyle As DataGridViewCellStyle
	Public ReadOnly Property MonthlyTotalRow As DataGridViewCellStyle
		Get
			Return _monthlyTotalRowCellStyle
		End Get
	End Property
	
	Private _allTotalRowCellStyle As DataGridViewCellStyle
	Public ReadOnly Property AllTotalRow As DataGridViewCellStyle
		Get
			Return _allTotalRowCellStyle
		End Get
	End Property
	
	Public Sub New(holidayRowColor As Color, monthlyTotalRowColor As Color, allTotalRowColor As Color)
		Me._holidayRowCellStyle = New DataGridViewCellStyle
		Me._holidayRowCellStyle.BackColor = holidayRowColor
		
		Me._monthlyTotalRowCellStyle = New DataGridViewCellStyle
		Me._monthlyTotalRowCellStyle.BackColor = monthlyTotalRowColor
		
		Me._allTotalRowCellStyle = New DataGridViewCellStyle
		Me._allTotalRowCellStyle.BackColor = allTotalRowColor
	End Sub
	
	Public Function SetDailyStyle(grid As DataGridView, beginRow As Integer, term As DateTerm) As Integer
'		If grid Is Nothing Then Throw New ArgumentNullException("gird is null")
'		
'		Dim lastRow As Integer = 
'			SetHeaderText(grid, beginRow, term.DailyLabelAndTermList,
'				Function(t)
'					Dim w As DayOfWeek = t.Min.DayOfWeek
'					Return w = DayOfWeek.Saturday OrElse w = DayOfWeek.Sunday
'				End Function,
'				Me._holidayRowCellStyle)
'
'		Return lastRow
	End Function
	
	Public Function SetWeeklyStyle(grid As DataGridView, beginRow As Integer, term As DateTerm) As Integer
'		If grid Is Nothing Then Throw New ArgumentNullException("gird is null")
'		Return SetHeaderText(grid, beginRow, term.WeeklyLabelAndTermList, Function(t) False, New DataGridViewCellStyle)
	End Function
	
	Public Function SetMonthlyStyle(grid As DataGridView, beginRow As Integer, term As DateTerm) As Integer
'		If grid Is Nothing Then Throw New ArgumentNullException("gird is null")
'		Return SetHeaderText(grid, beginRow, term.MonthlyLabelAndTermList, Function(t) True, Me._monthlyTotalRowCellStyle)
	End Function
	
	Private Function SetHeaderText()'grid As DataGridView, beginRow As Integer, labelAndTermList As List(Of LabelAndDateTerm), f As Func(Of DateTerm, Boolean), cellStyle As DataGridViewCellStyle) As Integer
'		Dim lastIndex As Integer = labelAndTermList.Count - 1
'		For i = 0 To lastIndex
'			Dim row As Integer = i + beginRow
'			If row < grid.Rows.Count Then
'				grid.Rows(row).HeaderCell.Value = labelAndTermList(i).Label
'				If f(labelAndTermList(i).Term) Then
'					grid.Rows(row).DefaultCellStyle = cellStyle
'				End If
'			End If
'		Next
'		
'		Dim lastRow As Integer = beginRow + lastIndex
'		If lastRow < grid.Rows.Count Then
'			Return lastRow
'		Else
'			Return grid.Rows.Count - 1
'		End If
	End Function
	
	Public Sub SetMonthlyTotalStyle(grid As DataGridView, row As Integer, headerText As String)
		SetTotalStyle(grid, row, headerText, Me._monthlyTotalRowCellStyle)
	End Sub
	
	Public Sub SetAllTotalStyle(grid As DataGridView, row As Integer, headerText As String)
		SetTotalStyle(grid, row, headerText, Me._allTotalRowCellStyle)
	End Sub
	
	Private Sub SetTotalStyle(grid As DataGridView, row As Integer, headerText As String, cellStyle As DataGridViewCellStyle)
		If row < grid.Rows.Count Then
			Dim r As DataGridViewRow = grid.Rows(row)
			r.HeaderCell.Value = headerText
			r.DefaultCellStyle = cellStyle
		End If		
	End Sub
	
	''' <summary>
	''' 指定したgridViewの行や列のヘッダーやセルの幅、高さを文字列の長さにあわせてリサイズする。
	''' </summary>
	Public Shared Sub AutoResizeAllCell(grid As DataGridView)
		' 行のヘッダーの幅を自動調整
		grid.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders)
		' 列のヘッダーの高さを自動調整
		grid.AutoResizeColumnHeadersHeight(DataGridViewColumnHeadersHeightSizeMode.AutoSize)
		' 全ての列の幅を自動調整
		grid.AutoResizeColumns()		
	End Sub
	
End Structure
