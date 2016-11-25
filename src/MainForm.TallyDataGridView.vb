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
  
  Private Sub InitTallyDataGridView()
    InitCboxTallyMonthly()
  End Sub
  
  ''' <summary>
  ''' 集計データの月のコンボボックスの要素を初期化する。
  ''' </summary>
  Sub InitCboxTallyMonthly()
    Dim monthly As List(Of DateTerm) = Me.dateTerm.MonthlyTerms()
    InitComboBox(
      Me.cboxTallyMonthly,
      monthly,
      GetType(DateTerm),
      Function(m) m.Label)
  End Sub
  
	Sub TabInTallyTab_SelectedIndexChanged(sender As Object, e As EventArgs)
		If initialized = False Then Return
		ShowTallyDataGridView()
	End Sub
	
	Public Sub ShowTallyDataGridView()
		Dim tabPage As TabPage = Me.tabInTallyTab.SelectedTab
		Dim pageName As String = tabPage.Text
		
		Dim grid As DataGridView = GetShowingDataGridViewInTallyDataPage()
		
    Try
      If pageName = "日" Then
        Dim term As DateTerm = DirectCast(Me.cboxTallyMonthly.SelectedValue, DateTerm)
        Dim table As DataTable = Me.userRecordManager.GetTotalOfAllUserDailyRecord(term, Me.chkBoxExcludeData.Checked)
        grid.DataSource = table
        
        SetColor(grid, term.BeginDate.Year, term.BeginDate.Month)
  		ElseIf pageName = "週"
        Dim table As DataTable = Me.userRecordManager.GetTotalOfAllUserWeeklyRecord(Me.dateTerm, Me.chkBoxExcludeData.Checked)
        grid.DataSource = table
        
        SetColorToOnlyTotalRow(grid)
  		ElseIf pageName = "月"
        Dim table As DataTable = Me.userRecordManager.GetTotalOfAllUserMonthlyRecord(Me.dateTerm, Me.chkBoxExcludeData.Checked)
        grid.DataSource = table
        
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
		Return TabPageUtils.GetDataGridView(tabPage)
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
