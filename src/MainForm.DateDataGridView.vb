'
' 日付: 2016/07/15
'
Imports System.Data
Imports Common.Account
Imports Common.Util
Imports Common.IO
Imports Common.Extensions

Public Partial Class MainForm
  
  Private Sub InitDateDataGridView()
    InitDateTimePicker()
    InitCboxWeekly()
    InitCboxMonthly()
  End Sub
  
	''' <summary>
	''' 日付データの日にち選択のコンボボックスの要素を初期化する。
	''' </summary>
	Private Sub InitDateTimePicker()
		Me.dateTimePickerInDatePage.MinDate = Me.dateTerm.BeginDate
		Me.dateTimePickerInDatePage.MaxDate = Me.dateTerm.EndDate
	End Sub
	
	''' <summary>
	''' 日付データの週のコンボボックスの要素を初期化する。
	''' </summary>
	Private Sub InitCboxWeekly()
	  ' 開始日がその月の中で何週目か取得する
	  Dim weekCountInMonth As Integer = DateUtils.GetWeekCountInMonth(Me.dateTerm.BeginDate, DayOfWeek.Saturday)
	  ' 期間を週単位で区切ったリストを取得する
    Dim weeklys As List(Of DateTerm) = _
      Me.dateTerm.WeeklyTerms(
        DayOfWeek.Saturday,
        Function(b, e)
          Dim str As String
          If b.Month = e.Month Then
            str = String.Format("{0}月第{1}週", b.Month, weekCountInMonth)
            weekCountInMonth += 1
          Else
            str = String.Format("{0}月第{1}週/{2}月第1週", b.Month, weekCountInMonth, e.Month)
            weekCountInMonth = 2
          End If
          
          Return str
        End Function)
    
    ' 週単位でコンボボックスにセットする
    InitComboBox(
      Me.cboxWeekly,
      weeklys,
      GetType(DateTerm),
      Function(w) w.Label)
  End Sub
  
  ''' <summary>
  ''' 日付データの月のコンボボックスの要素を初期化する。
  ''' </summary>
  Sub InitCboxMonthly()
    Dim monthly As List(Of DateTerm) = Me.dateTerm.MonthlyTerms()
    InitComboBox(
      Me.cboxMonthly,
      monthly,
      GetType(DateTerm),
      Function(m) m.Label)
  End Sub
  
  
	Sub TabInDateTab_SelectedIndexChanged(sender As Object, e As EventArgs)
		If initialized = False Then Return
		ShowDateDataGridView()
	End Sub

	''' <summary>
	''' 指定したページのGridViewを表示する。
	''' </summary>
	Private Sub ShowDateDataGridView()
		Dim tabPage As TabPage = Me.tabInDateTab.SelectedTab
'		Dim isDepending As Boolean = Me.chkBoxExcludeData.Checked
'		Dim userDataList As List(Of UserData) = Me.userRecordManager.UserDataList
'		
'		If userDataList.Count = 0 Then
'			Return
'		End If
'		
		Dim term As New DateTerm(#01/01/1900#, #01/01/1900#)
		If tabPage.Text = "日" Then
			Dim datePicker As DateTimePicker = TabPageUtils.GetDateTimePicker(tabPage)
			If datePicker IsNot Nothing Then 
				term = New DateTerm(datePicker.Value, datePicker.Value, datePicker.Value.Day.ToString & "日")
			End If
			
			Log.out("picked date: " & term.ToString)
		ElseIf tabPage.Text = "週" OrElse tabPage.Text = "月"
			Dim cbox As ComboBox = TabPageUtils.GetComboBox(tabPage)
			If cbox IsNot Nothing Then
				term = DirectCast(cbox.SelectedValue, DateTerm)
			End If
		ElseIf tabPage.Text = "合計"
			term = Me.dateTerm
		End If		
		
		Try
  		Dim grid As DataGridView = GetShowingDataGridViewInDateDataPage()
  		
'  		If grid IsNot Nothing AndAlso term.EndDate <> #01/01/1900# Then
'  		  Dim table As DataTable = Me.userRecordManager.GetTallyRecordOfEachUser(term)
'  		  grid.DataSource = table
'  		End If
      If grid IsNot Nothing AndAlso term.EndDate <> #01/01/1900# Then
        Dim table As DataTable = Me.userRecordManager.GetTallyRecordOfEachUser(term, Me.chkBoxExcludeData.Checked)
'        Dim totalRow As DataRow = table.NewRow
'        totalRow(UserRecord.NAME_COL_NAME) = "合計"
'        For Each row As DataRow In table.Rows
'          totalRow.PlusByDouble(row)
'        Next
'        table.Rows.Add(totalRow)
        grid.DataSource = table
        HoldFirstColumn(grid)
        SetViewSize(grid, Me.userRecordManager.GetUserRecordColumnsInfo)
        SetColorToOnlyTotalRow(grid)
      End If
		Catch ex As Exception
		  MsgBox.ShowError(ex)
		End Try
		
'			Dim newTable As DataTable = userDataList(0).Record.CreateTable
'			' 各ユーザごとの行を取得
'			userDataList.ForEach(Sub(data) data.Record.AddTotalRowToTable(term, isDepending, newTable))
'			' 合計値を最後の行に追加
'			userDataList(0).Record.AddTotalRowToTable(newTable, isDepending)
'			
'			grid.DataSource = newTable
'			
'			' グリッドの表示スタイルをセット
'			SetGridStyleToDateDataGridView(grid, userDataList)
'			' セルの幅、高さを自動調整
'			MyGridViewCellStyles.AutoResizeAllCell(grid)
'		End If
	End Sub
	
	''' <summary>
	''' 現在表示されているページのDataGridViewを返す。
	''' </summary>
	Function GetShowingDataGridViewInDateDataPage() As DataGridView
	  Dim tabPage As TabPage = Me.tabInDateTab.SelectedTab
		Return TabPageUtils.GetDataGridView(tabPage)
	End Function
	
	''' <summary>
	''' グリッドの表示スタイルをセットする。
	''' </summary>
	Private Sub SetGridStyleToDateDataGridView(grid As DataGridView)', userDataList As List(Of UserData))
'		For i = 0 To userDataList.Count - 1
'			grid.Rows(i).HeaderCell.Value = userDataList(i).Name
'		Next
'		grid.Rows(grid.Rows.Count - 1).HeaderCell.Value = "合計"
'		
'		grid.Rows(grid.Rows.Count - 1).DefaultCellStyle = Me.gridViewCellStyles.MonthlyTotalRow
	End Sub
	

End Class
