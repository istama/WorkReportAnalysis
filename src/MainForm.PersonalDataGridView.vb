'
' 日付: 2016/07/15
'
Imports System.Data

Imports Common.account
Imports Common.Util
Imports Common.Extensions

''' <summary>
''' メインフォーム内の個人ページに関する処理をまとめた部分フォーム。
''' </summary>
Public Partial Class MainForm
	Private Const PAGE_NAME_TOTAL As String = "集計"
	
	Private pageNameAndTermDictionary As IDictionary(Of String, DateTerm)
	
	Private Sub InitPersonalDataGridView()
		Me.pageNameAndTermDictionary = New Dictionary(Of String, dateTerm)
		Me.InitTabPageInPersonalTab()
	End Sub
	
	''' <summary>
	''' タブに月ごとのページが用意されるよう初期化する。
	''' </summary>
	Private Sub InitTabPageInPersonalTab()
		Dim isPageOne As Boolean = True ' 1ページ目かどうか

    Dim monthly As List(Of dateTerm) = 
      Me.dateTerm.MonthlyTerms(Function(b, e) String.Format("{0}月分", b.Month.ToString))

	 	  ' 期間を月単位で区切ってタブページを設定
		  monthly.ForEach(
  		  Sub(term)
    			Dim pageName As String = Me.excelProperties.SheetName(term.BeginDate.Month)
    			
    			' １ページ目は既に用意されているのでページを作成しない
    			If isPageOne Then
    			  'AddHandler Me.gridMonth10InPersonal.ColumnHeaderMouseClick, AddressOf SortDataGridView
    			  
    				Me.tabInPersonalTab.TabPages.Item(0).Text = pageName
    				isPageOne = False				
    			Else
    				Me.tabInPersonalTab.TabPages.Add(CreateTabPage(pageName))
    			End If
    			
    			' ページ名と日付をひもつける
    			Me.pageNameAndTermDictionary.Add(pageName, term)
        End Sub)
  		
    Me.tabInPersonalTab.TabPages.Add(CreateTabPage(PAGE_NAME_TOTAL))
	End Sub
	
	''' <summary>
	''' タブのページを生成する。
	''' </summary>
	Private Function CreateTabPage(pageName As String) As TabPage
	  ' 新しいタブページを作成
		Dim page As New TabPage
		page.Name    = "grid" & pageName
		page.Text    = pageName
		page.Padding = New Padding(3)
		page.Margin  = New Padding(3)
		
		' 作成したタブページ内にDataGridViewのコントロールを追加する
		Dim grid As New DataGridView
		grid.Location = New Point(3, 3)
		grid.Margin   = New Padding(3)
		grid.Dock     = DockStyle.Fill
		grid.ScrollBars = ScrollBars.Both
		grid.AllowUserToAddRows = False
		
		'AddHandler grid.ColumnHeaderMouseClick, AddressOf SortDataGridView
		
		page.Controls.Add(grid)
		
		Return page
	End Function
	
	''' <summary>
	''' 個人ページ内のページを変更したときに発生するイベント。
	''' </summary>
	Sub TabInPersonalTab_SelectedIndexChanged(sender As Object, e As EventArgs)
		If initialized = False Then	Return
		ShowPersonalDataGridView()
	End Sub
	
	''' <summary>
	''' 指定したページのGridViewを表示する。
	''' </summary>
	Private Sub ShowPersonalDataGridView()
		Dim tabPage As TabPage = Me.tabInPersonalTab.SelectedTab
		Dim grid As DataGridView = TabPageUtils.GetDataGridView(tabPage)
		Dim pageName As String = tabPage.Text
		
		' 現在ページの月を取得する
		Dim month As Integer = -1
		For m As Integer = 1 To 12
		  If pageName = Me.excelProperties.SheetName(m) Then
		    month = m
		    Exit For
		  End If
		Next
		
'		
'		Dim isDepending As Boolean = Me.chkBoxExcludeData.Checked
'		
    Try
  		' ComboBox.DataSourceで要素を格納した場合は、ComboBox.SelectedItemではなく、
  		' ComboBox.SelectedValueで値を取得する。
  		Dim userInfo As UserInfo = DirectCast(Me.cboxUserName.SelectedValue, UserInfo)
  		If userInfo IsNot Nothing Then
  		  If month >= 1 AndAlso month <= 12 Then 
  		    Dim table As DataTable = Me.userRecordManager.GetDailyRecordLabelingDate(userInfo, Me.dateTerm.BeginDate.Year, month)
'          Dim totalRow As DataRow = table.NewRow
'          totalRow(UserRecord.DATE_COL_NAME) = "合計"
'          For Each row As DataRow In table.Rows
'            totalRow.PlusByDouble(row)
'          Next
'          table.Rows.Add(totalRow)
          
          grid.DataSource = table
          SetViewSize(grid, Me.userRecordManager.GetUserRecordColumnsInfo)
          SetColor(grid, Me.dateTerm.BeginDate.Year, month)
  		  Else
  		    ' 集計ページを表示する
  		    Dim table As DataTable = Me.userRecordManager.GetSumRecord(userInfo)
  		    grid.DataSource = table
  		    SetViewSize(grid, Me.userRecordManager.GetUserRecordColumnsInfo)
  		    SetColorToOnlyTotalRow(grid)
  		  End If
  		End If
    Catch ex As Exception
      MsgBox.ShowError(ex)
    End Try
		
		
'		Dim record As UserRecord = Me.userRecordManager.GetUserRecord(userInfo)
'		Dim gridTable As GridTable = New GridTable(record, Me.dateTerm)
'		
'		If record.HasTable(pageName) Then
'			' 月のレコードを表示する
'			grid.DataSource = gridTable.GetMonthlyTable(pageName, isDepending)
'			Dim lastRow As Integer = Me.gridViewCellStyles.SetDailyStyle(grid, 0, Me.pageNameAndTermDictionary(pageName))
'			Me.gridViewCellStyles.SetMonthlyTotalStyle(grid, lastRow + 1, "合計")
'		ElseIf pageName = PAGE_NAME_TOTAL
'			' 集計レコードを表示する
'			grid.DataSource = gridTable.GetTotalTable(isDepending)
'			Dim lastRow As Integer = Me.gridViewCellStyles.SetWeeklyStyle(grid, 0, Me.dateTerm)
'			lastRow = Me.gridViewCellStyles.SetMonthlyStyle(grid, lastRow + 2, Me.dateTerm)
'			Me.gridViewCellStyles.SetAllTotalStyle(grid, lastRow + 2, "合計")
'		Else
'			Return
'		End If
'		'grid.Columns(0).SortMode = DataGridViewColumnSortMode.Automatic
'		AddHandler grid.ColumnHeaderMouseClick, AddressOf DataGridView_ColumnHeaderMouseClick
'		'AddHandler grid.SortCompare, AddressOf DataGridView1_SortCompare
'		' セルの幅、高さを自動調整
'		MyGridViewCellStyles.AutoResizeAllCell(grid)
	End Sub
	
	Private Sub DataGridView_ColumnHeaderMouseClick(sender As Object, e As EventArgs)
'		Dim grid As DataGridView = DirectCast(sender, DataGridView)
'		Dim args As DataGridViewCellMouseEventArgs = DirectCast(e, DataGridViewCellMouseEventArgs)
'		
'		Dim dt As DataTable = CType(grid.DataSource, DataTable)
'		
'		'DataViewを取得
'		Dim dv As DataView = dt.DefaultView
'		Dim colName As String = grid.Columns(args.ColumnIndex).Name
'		
'		dv.Sort = colName & " ASC"
	End Sub
	
'	Private Sub DataGridView1_SortCompare(sender As Object, e As EventArgs)
'		'Handles DataGridView1.SortCompare
'		Dim args As DataGridViewSortCompareEventArgs = DirectCast(e, DataGridViewSortCompareEventArgs)
'		
'    '指定されたセルの値を文字列として取得する
'    Dim v1 As Double = 0.0
'    If args.CellValue1 IsNot Nothing Then
'			Double.TryParse(args.CellValue1.ToString, v1) 
'    End If
'    
'    Dim v2 As Double = 0.0
'    If args.CellValue2 IsNot Nothing Then
'    	Double.TryParse(args.CellValue2.ToString, v2)
'    End If
'
'    '結果を代入
'    args.SortResult = CType(v1 - v2, Integer)
'    '処理したことを知らせる
'    args.Handled = True
'	End Sub
End Class

'Public Class CustomComparer
'	Implements IComparer
'	Private colIdx As Integer
'	Private order As SortOrder
'	'Private comparer As Comparer
'	
'	Public Sub New(ByVal order As SortOrder, colIdx As Integer)
'		Me.order = order
'		Me.colIdx = colIdx
'		'Me.comparer = New Comparer(System.Globalization.CultureInfo.CurrentCulture)
'	End Sub
'	
'	'並び替え方を定義する
'	Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
'		Implements System.Collections.IComparer.Compare
'		
'		Dim rowx As DataGridViewRow = CType(x, DataGridViewRow)
'		Dim rowy As DataGridViewRow = CType(y, DataGridViewRow)
'		
'		Dim v1 As Double = 0.0
'		If rowx.Cells(Me.colIdx).Value IsNot Nothing Then
'			Double.TryParse(rowx.Cells(Me.colIdx).Value.ToString, v1)
'		End If
'		
'		Dim v2 As Double = 0.0
'		If rowy.Cells(Me.colIdx).Value IsNot Nothing Then
'			Double.TryParse(rowy.Cells(Me.colIdx).Value.ToString, v2)
'		End If		
'		
'		Dim result As Integer
'		
'		If v1 <> 0.0 Then
'			If v2 <> 0.0 Then
'				result = DirectCast(IIf(v1 < v2, -1, IIf(v1 = v2, 0, 1)), Integer)
'				' 降順の場合
'				If order = SortOrder.Descending Then
'					result *= -1
'				End If
'			Else
'				result = -1
'			End If
'		ElseIf v2 <> 0.0
'			result = 1
'		Else
'			result = -1
'		End If
'		
'		'結果を返す
'		Return result
'	End Function
'End Class