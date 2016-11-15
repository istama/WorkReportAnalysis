'
' 日付: 2016/11/15
'
Imports Common.Util

Public Partial Class MainForm
  ''' DataGridViewにおける各作業項目の列のサイズ
  Private ReadOnly WORKITEM_COLUMN_SIZE As Integer = 50
  ''' DataGridViewにおける土日の行の色  
  Private ReadOnly HOLYDAY_ROW_COLOR As Color = Color.LightPink
  ''' DataGridViewにおける合計行の色
  Private ReadOnly TOTAL_ROW_COLOR   As Color = Color.LightGreen
  
  ''' <summary>
  ''' DataGridViewの各セルの表示サイズを設定する。
  ''' </summary>
  Public Sub SetViewSize(view As DataGridView, recordColumnsInfo As UserRecordColumnsInfo)
    If view Is Nothing Then Throw New NullReferenceException("view is null")
    
    ' 列ヘッダーのサイズが要素に合わせて自動に設定されるようにする
    view.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
    ' 各列のサイズが要素に合わせて自動に設定されるようにする
    view.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
    
    For Each col In recordColumnsInfo.WorkItemList
      If Not String.IsNullOrWhiteSpace(col.WorkCountColName) Then
        view.Columns(col.WorkCountColName).Width = WORKITEM_COLUMN_SIZE
      End If
      
      If Not String.IsNullOrWhiteSpace(col.WorkTimeColName) Then
        view.Columns(col.WorkTimeColName).Width = WORKITEM_COLUMN_SIZE
      End If
      
      If Not String.IsNullOrWhiteSpace(col.WorkProductivityColName) Then
        view.Columns(col.WorkProductivityColName).Width = WORKITEM_COLUMN_SIZE
      End If
    Next
    
    view.Columns(recordColumnsInfo.noteColName).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
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
End Class
