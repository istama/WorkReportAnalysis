'
' 日付: 2016/07/19
'
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Data

Imports Common.account
Imports Common.Util
Imports Common.Threading

Public Partial Class MainForm
  
  ''' <summary>
  ''' タブページを初期化する。
  ''' </summary>
  Private Sub InitTallyDataGridView()
    InitCboxTallyMonthly()
    InitTabPageInTallyTab()
  End Sub
  
  ''' <summary>
  ''' 集計データの月のコンボボックスの要素を初期化する。
  ''' </summary>
  Sub InitCboxTallyMonthly()
    ' 期間を月単位で区切ったリストを取得する
    Dim monthly As IEnumerable(Of DateTerm) =
      Me.dateTerm.MonthlyTerms(Function(begin, _end) begin.Month.ToString & "月")
    
    ' 月単位でコンボボックスにセットする
    InitComboBox(
      Me.cboxTallyMonthly,
      monthly,
      GetType(DateTerm),
      Function(m) m.Label)
  End Sub
  
  ''' <summary>
  ''' タブページの要素を初期化する。
  ''' </summary>
  Sub InitTabPageInTallyTab()
    ' GridViewの列名をクリックされたときに実行されるハンドラを登録する
    AddHandler Me.gridDailyInTally.ColumnHeaderMouseClick, AddressOf SortDataGridView
    AddHandler Me.gridWeeklyInTally.ColumnHeaderMouseClick, AddressOf SortDataGridView
    AddHandler Me.gridMonthlyInTally.ColumnHeaderMouseClick, AddressOf SortDataGridView
  End Sub
  
	Sub TabInTallyTab_SelectedIndexChanged(sender As Object, e As EventArgs)
		If initialized = False Then Return
		ShowTallyDataGridView()
	End Sub
	
	Public Sub ShowTallyDataGridView()
		Dim pageName As String       = Me.tabInTallyTab.SelectedTab.Text
		Dim grid     As DataGridView = GetShowingDataGridViewInTallyDataPage()
		
    Try
      If pageName = "日" Then
        Dim term As DateTerm = DirectCast(Me.cboxTallyMonthly.SelectedValue, DateTerm)
        grid.DataSource = Me.userRecordManager.GetDailyTotalRecord(term, Me.chkBoxExcludeData.Checked)
        
        SetColor(grid, term.BeginDate.Year, term.BeginDate.Month)
  		ElseIf pageName = "週"
        grid.DataSource = Me.userRecordManager.GetWeeklyTotalRecord(Me.dateTerm, Me.chkBoxExcludeData.Checked)
        
        SetColorToOnlyTotalRow(grid)
  		ElseIf pageName = "月"
        grid.DataSource = Me.userRecordManager.GetMonthlyTotalRecord(Me.dateTerm, Me.chkBoxExcludeData.Checked)
        
        SetColorToOnlyTotalRow(grid)
  		Else
  			Return
  		End If
  		
  		HoldFirstColumn(grid)
  		SetViewSize(grid, Me.userRecordManager.GetUserRecordColumnsInfo)
    Catch ex As Exception
      MsgBox.ShowError(ex)      
    End Try
	End Sub
		
	''' <summary>
	''' 現在表示されているページのDataGridViewを返す。
	''' </summary>
	Function GetShowingDataGridViewInTallyDataPage() As DataGridView
	  Dim tabPage As TabPage = Me.tabInTallyTab.SelectedTab
		Return GetChildControl(Of DataGridView)(tabPage)
	End Function
	
	''' <summary>
	''' 現在表示されているデータの名前を返す。
	''' </summary>
	Function GetShowingDataNameInTallyDatePage() As String
	  Dim pageName As String =  Me.tabInTallyTab.SelectedTab.Text
	  If pageName = "日" Then
	    Dim term As DateTerm = DirectCast(Me.cboxTallyMonthly.SelectedValue, DateTerm)
	    Return "集計データ_" & term.BeginDate.Month.ToString & "月"
	  ElseIf pageName = "週"
	    Return "集計データ_" & "週"
	  ElseIf pageName = "月"
	    Return "集計データ_" & "月"
	  Else
	    Return String.Empty
	  End If
	End Function
	
End Class
