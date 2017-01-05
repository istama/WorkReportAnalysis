'
' 日付: 2016/11/15
'
Imports System.Linq
Imports System.Data
Imports Common.Extensions
Imports Common.Util
Imports Common.IO

Public Partial Class MainForm
  ''' DataGridViewにおける各作業項目の列のサイズ
  Private ReadOnly WORKITEM_COLUMN_SIZE As Integer = 50
  ''' DataGridViewにおける土日の行の色  
  Private ReadOnly HOLYDAY_ROW_COLOR As Color = Color.LightPink
  ''' DataGridViewにおける合計行の色
  Private ReadOnly TOTAL_ROW_COLOR   As Color = Color.LightGreen
  
  ''' DataTableをソートするために使用する。
  Private ReadOnly dataTableCompare As New DataTableCompare()
  
  ''' <summary>
  ''' 現在表示されているDataGridViewを取得する。
  ''' </summary>
  Function GetShowingDataGridView() As DataGridView
		Dim pageName = Me.tabRoot.SelectedTab.Text
		If pageName = TABPAGE_NAME_PERSONAL Then
			Return Me.GetDataGridView(Me.tabInPersonalTab.SelectedTab)
		ElseIf pageName = TABPAGE_NAME_DATE
			Return Me.GetDataGridView(Me.tabInDateTab.SelectedTab)
		ElseIf pageName = TABPAGE_NAME_TOTAL
		  Return Me.GetDataGridView(Me.tabInTallyTab.SelectedTab)
		Else
		  Return Nothing
		End If		
  End Function
  
  ''' <summary>
  ''' 現在表示されているDataGridViewを取得する。
  ''' </summary>
  Function GetShowingDataName() As String
		Dim pageName = Me.tabRoot.SelectedTab.Text
		If pageName = TABPAGE_NAME_PERSONAL Then
			Return GetShowingDataNameInPersonalDataPage()
		ElseIf pageName = TABPAGE_NAME_DATE
			Return GetShowingDataNameInDateDataPage()
		ElseIf pageName = TABPAGE_NAME_TOTAL
		  Return GetShowingDataNameInTallyDatePage()
		Else
		  Return Nothing
		End If		
  End Function
  
  ''' <summary>
  ''' 最初の列を横スクロール時に移動しないよう固定する。
  ''' </summary>
  Public Sub HoldFirstColumn(view As DataGridView)
    view.Columns(0).Frozen = True    
  End Sub
  
  ''' <summary>
  ''' DataGridViewの各セルの表示サイズを設定する。
  ''' </summary>
  Public Sub SetViewSize(view As DataGridView, recordColumnsInfo As UserRecordColumnsInfo)
    If view Is Nothing Then Throw New NullReferenceException("view is null")
    
    ' 列ヘッダーのサイズが要素に合わせて自動に設定されるようにする
    view.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
    ' 各列のサイズが要素に合わせて自動に設定されるようにする
    view.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
    
    For Each item In recordColumnsInfo.WorkItems
      If Not String.IsNullOrWhiteSpace(item.WorkCountColInfo.name) Then
        view.Columns(item.WorkCountColInfo.name).Width = WORKITEM_COLUMN_SIZE
      End If
      
      If Not String.IsNullOrWhiteSpace(item.WorkTimeColInfo.name) Then
        view.Columns(item.WorkTimeColInfo.name).Width = WORKITEM_COLUMN_SIZE
      End If
      
      If Not String.IsNullOrWhiteSpace(item.WorkProductivityColInfo.name) Then
        view.Columns(item.WorkProductivityColInfo.name).Width = WORKITEM_COLUMN_SIZE
      End If
    Next
    
    view.Columns(recordColumnsInfo.noteColInfo.Name).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
  End Sub
  
  ''' <summary>
  ''' DataGridViewの行に色をつける。
  ''' 色が付く行は、土曜、日曜、合計行。
  ''' </summary>
  Public Sub SetColor(view As DataGridView, year As Integer, month As Integer)
    ' １ヶ月の中でもっとも早い土曜日と日曜日を取得する
    Dim saturday As DateTime = DateUtils.GetDateOfNextWeekDay(New DateTime(year, month, 1), DayOfWeek.Saturday) 
    Dim sunday   As DateTime = DateUtils.GetDateOfNextWeekDay(New DateTime(year, month, 1), DayOfWeek.Sunday)
    
    ' 土曜日からループ文で色を付けていくので、日曜日の方が早い場合、前もって色をつけておく
    If sunday < saturday Then
      view.Rows(sunday.Day - 1).DefaultCellStyle.BackColor = HOLYDAY_ROW_COLOR
    End If
      
    Dim holyDay As DateTime = saturday
    
    Do
      ' 土曜日の行に色をつける
      view.Rows(holyDay.Day - 1).DefaultCellStyle.BackColor = HOLYDAY_ROW_COLOR
      holyDay = holyDay.AddDays(1)
      
      ' 次の日曜が月をまたがない場合、日曜の行に色をつける
      If holyDay.Month = month Then
        view.Rows(holyDay.Day - 1).DefaultCellStyle.BackColor = HOLYDAY_ROW_COLOR
        holyDay = holyDay.AddDays(6)
      End If
    Loop While holyDay.Month = month 
    
    ' 合計行に色をつける。
    SetColorToOnlyTotalRow(view)
  End Sub
  
  ''' <summary>
  ''' 合計行に色をつける。
  ''' </summary>
  Public Sub SetColorToOnlyTotalRow(view As DataGridView)
    view.Rows(view.Rows.Count - 1).DefaultCellStyle.BackColor = TOTAL_ROW_COLOR
  End Sub
  
  ''' <summary>
  ''' DataGridViewをソートする。
  ''' </summary>
  Public Sub SortDataGridView(sender As Object, e As DataGridViewCellMouseEventArgs)
    Dim grid  As DataGridView = DirectCast(sender, DataGridView)
    Dim table As DataTable    = DirectCast(grid.DataSource, DataTable)
    
    Dim list As New List(Of DataRow)
    For Each dataRow As DataRow In table.Rows
      list.Add(dataRow)
    Next
    
    ' 指定した列で行をソートする
    list.Sort(Me.dataTableCompare.GetDataRowCompare(table, e.ColumnIndex))
    
    grid.DataSource = list.CopyToDataTable()
  End Sub
  
	Public Function GetDataGridView(tabPage As TabPage) As DataGridView
		Dim grid As DataGridView = Nothing
		For i = 0 To tabPage.Controls.Count
			grid = TryCast(tabPage.Controls.Item(i), DataGridView)
			If grid IsNot Nothing Then
				Exit For
			End If
		Next
		
		Return grid
	End Function
	
	Public Function GetComboBox(tabPage As TabPage) As ComboBox
		Dim cbox As ComboBox = Nothing
		For i = 0 To tabPage.Controls.Count
			cbox = TryCast(tabPage.Controls.Item(i), ComboBox)
			If cbox IsNot Nothing Then
				Exit For
			End If
		Next
		
		Return cbox		
	End Function
	
	Public Function GetDateTimePicker(tabPage As TabPage) As DateTimePicker
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

