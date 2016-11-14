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
		Dim grid As DataGridView = TabPageUtils.GetDataGridView(tabPage)
		
'		Dim userDataList As List(Of UserData) = Me.userRecordManager.UserDataList
'		Dim newTable As DataTable = userDataList(0).Record.CreateTable
'		Dim rowHeaderList As New List(Of String)
'		
'		Dim watch As New Stopwatch
'		watch.Start
    Try
      If pageName = "日" Then
        Dim term As DateTerm = DirectCast(Me.cboxTallyMonthly.SelectedValue, DateTerm)
        Dim table As DataTable = Me.userRecordManager.GetTotalOfAllUserDailyRecord(term)
        grid.DataSource = table
  '			LoadTallyTable(termList, userDataList, newTable, Me.chkBoxExcludeData.Checked)
  '			grid.DataSource = newTable
  '			Dim lastRow As Integer = Me.gridViewCellStyles.SetDailyStyle(grid, 0, dTerm)
  '			Me.gridViewCellStyles.SetMonthlyTotalStyle(grid, lastRow + 1, "合計")
  		ElseIf pageName = "週"
        Dim table As DataTable = Me.userRecordManager.GetTotalOfAllUserWeeklyRecord(Me.dateTerm)
        grid.DataSource = table
  '			Dim termList As List(Of DateTerm) = Me.dateTerm.WeeklyTerms
  '			LoadTallyTable(termList, userDataList, newTable, Me.chkBoxExcludeData.Checked)
  '			grid.DataSource = newTable
  '			Dim lastRow As Integer = Me.gridViewCellStyles.SetWeeklyStyle(grid, 0, Me.dateTerm)
  '			Me.gridViewCellStyles.SetAllTotalStyle(grid, lastRow + 1, "合計")			
  		ElseIf pageName = "月"
        Dim table As DataTable = Me.userRecordManager.GetTotalOfAllUserMonthlyRecord(Me.dateTerm)
        grid.DataSource = table
  '			Dim termList As List(Of DateTerm) = Me.dateTerm.MonthlyTerms
  '			LoadTallyTable(termList, userDataList, newTable, Me.chkBoxExcludeData.Checked)
  '			grid.DataSource = newTable
  '			Dim lastRow As Integer = Me.gridViewCellStyles.SetMonthlyStyle(grid, 0, Me.dateTerm)
  '			Me.gridViewCellStyles.SetAllTotalStyle(grid, lastRow + 1, "合計")
  		Else
  			Return
  		End If
    Catch ex As Exception
      MsgBox.ShowError(ex)      
    End Try
'		watch.Stop
'		MsgBox.Show(watch.Elapsed.ToString, "")
'		
'		' セルの幅、高さを自動調整
'		MyGridViewCellStyles.AutoResizeAllCell(grid)
	End Sub
	
	Private Sub LoadTallyTable()'(termList As List(Of DateTerm), userDataList As List(Of UserData), resultTable As DataTable, isDependent As Boolean)
'		If termList     Is Nothing Then Throw New ArgumentNullException("termList is null")
'		If userDataList Is Nothing Then Throw New ArgumentNullException("userDataList is null")
'		If resultTable  Is Nothing Then Throw New ArgumentNullException("resultTable is null")
'		
'		If userDataList.Count = 0 Then
'			Return
'		End If
'		
''		Dim tmpTable As DataTable = userDataList(0).Record.CreateTable
''		Dim totalRowList As List(Of DataRow) =
''			MultiTask.Run(Of DateTerm, DataRow)(
''				termList,
''				Function(term)
''					Return Function(obj)
''						Dim usersTable As DataTable = userDataList(0).Record.CreateTable
''						' 各ユーザの、指定した期間中の合計値を算出しテーブルに追加する
''						MultiTask.Run(Of UserData)(
''							userDataList,
''							Function(data) Sub(o) data.Record.AddTotalRowToTable(term, isDependent, usersTable),
''							Nothing)
''						' 上記の各ユーザの合計値の合計値を算出しテーブルに追加する
''						Dim newRow As DataRow = tmpTable.NewRow
''						userDataList(0).Record.TotalRow(usersTable, isDependent, newRow)
''						
''						Return newRow
''					End Function
''				End Function,
''				Nothing)
''		
''		totalRowList.ForEach(
''			Sub(row)
''				Dim newRow As DataRow = resultTable.NewRow
''				newRow.ItemArray = row.ItemArray
''				resultTable.Rows.Add(newRow)
''			End Sub)
'				
'		Dim usersTable As DataTable = userDataList(0).Record.CreateTable
'		termList.ForEach(
'			Sub(term)
'				' 各ユーザの、指定した期間中の合計値を算出しテーブルに追加する
'				MultiTask.Run(Of UserData)(
'					userDataList,
'					Function(data) Sub(o) data.Record.AddTotalRowToTable(term, isDependent, usersTable),
'					Nothing)
'				
'				' 上記の各ユーザの合計値の合計値を算出しテーブルに追加する
'				userDataList(0).Record.AddTotalRowToTable(usersTable, isDependent, resultTable)
'				usersTable.Clear
'			End Sub)
'		
'		' すべての合計値を算出しテーブルに追加する
'		userDataList(0).Record.AddTotalRowToTable(resultTable, isDependent)
	End Sub	
End Class
