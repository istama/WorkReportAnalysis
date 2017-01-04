'
' 日付: 2016/07/15
'
Imports System.Data
Imports Common.Account
Imports Common.Util
Imports Common.IO
Imports Common.Extensions

Public Partial Class MainForm
  
  ''' <summary>
  ''' タブページを初期化する。
  ''' </summary>
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
	  ' 週はじめと週末の日付を取得して、月の第何週かを表す文字列を返す関数
	  Dim funcToGetWeekCount As Func(Of DateTime, DateTime, String) =
	    Function(weekStart, weekEnd)
	      Dim cnt As Integer = DateUtils.GetWeekCountInMonth(weekStart, DayOfWeek.Saturday)
	      Dim str As String  = String.Format("{0}月第{1}週", weekStart.Month, cnt)
	      
        If weekStart.Month <> weekEnd.Month Then
          str = str & String.Format("/{0}月第1週", weekEnd.Month)
        End If
        
        Return str
      End Function
	  
	  
	  ' 期間を週単位で区切ったリストを取得する
    Dim weeklys As IEnumerable(Of DateTerm) = _
      Me.dateTerm.WeeklyTerms(DayOfWeek.Saturday, funcToGetWeekCount)
    
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
    ' 期間を月単位で区切ったリストを取得する
    Dim monthly As IEnumerable(Of DateTerm) =
      Me.dateTerm.MonthlyTerms(Function(begin, _end) begin.Month.ToString & "月")
    
    ' 月単位でコンボボックスにセットする
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
    Dim term As DateTerm = GetShowingDataDateTerm()
		
		Try
  		Dim grid As DataGridView = GetShowingDataGridViewInDateDataPage()
  		
  		If grid IsNot Nothing AndAlso term.EndDate <> #01/01/1900# Then
        grid.DataSource = Me.userRecordManager.GetAllUserSumRecord(term, Me.chkBoxExcludeData.Checked)
        
        HoldFirstColumn(grid) ' ビューの横スクロール時に１列目を固定して表示するようにする
        SetViewSize(grid, Me.userRecordManager.GetUserRecordColumnsInfo)
        SetColorToOnlyTotalRow(grid)
      End If
		Catch ex As Exception
		  MsgBox.ShowError(ex)
		End Try
		
	End Sub
	
	''' <summary>
	''' 現在表示されているページのDataGridViewを返す。
	''' </summary>
	Function GetShowingDataGridViewInDateDataPage() As DataGridView
	  Dim tabPage As TabPage = Me.tabInDateTab.SelectedTab
		Return GetDataGridView(tabPage)
	End Function
	
	''' <summary>
	''' 現在表示されているデータの名前を取得する
	''' </summary>
	''' <returns></returns>
	Function GetShowingDataNameInDateDataPage() As String
	  Dim term As DateTerm = GetShowingDataDateTerm()
	  
	  Dim pageName As String = Me.tabInDateTab.SelectedTab.Text
	  If pageName = "日" Then
	    Dim d As DateTime = term.BeginDate
	    Return "日付データ_" & d.ToString("yyMMdd")
	  ElseIf pageName = "週" OrElse pageName = "月"
	    Return "日付データ_" & term.Label
	  Else
	    Return "各ユーザ合計"	    
	  End If
	End Function
	
	''' <summary>
	''' 現在表示されているデータの期間を取得する。
	''' </summary>
	Private Function GetShowingDataDateTerm() As DateTerm
	  Dim tabPage As TabPage = Me.tabInDateTab.SelectedTab
	  Dim pageName As String = tabPage.Text
	  If pageName = "日" Then
	    Dim d As DateTime = Me.dateTimePickerInDatePage.Value
	    Return New DateTerm(d, d, d.Day.ToString & "日")
	  ElseIf pageName = "週" OrElse pageName = "月"
	    Dim cbox As ComboBox = GetComboBox(tabPage)
			Return DirectCast(cbox.SelectedValue, DateTerm)
	  Else
	    Return Me.dateTerm
	  End If	  
	End Function
	
End Class
