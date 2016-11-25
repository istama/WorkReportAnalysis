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
    ' 現在ページのDataGridViewを取得する。
    Dim grid As DataGridView = GetShowingDataGridViewInPersonalDataPage()
		' 現在ページの月を取得する
		Dim month As Integer = GetSelectedPageMonthInPersonalDataPage()
		
		Try
  		' ComboBox.DataSourceで要素を格納した場合は、ComboBox.SelectedItemではなく、
  		' ComboBox.SelectedValueで値を取得する。
  		Dim userInfo As UserInfo = GetSelectedUserInfo()
  		If userInfo IsNot Nothing Then
  		  If month >= 1 AndAlso month <= 12 Then 
  		    Dim table As DataTable =
  		      Me.userRecordManager.GetDailyRecordLabelingDate(userInfo, Me.dateTerm.BeginDate.Year, month)
          
          grid.DataSource = table
          HoldFirstColumn(grid)
          SetViewSize(grid, Me.userRecordManager.GetUserRecordColumnsInfo)
          SetColor(grid, Me.dateTerm.BeginDate.Year, month)
  		  Else
  		    ' 集計ページを表示する
  		    Dim table As DataTable = Me.userRecordManager.GetSumRecord(userInfo, Me.chkBoxExcludeData.Checked)
  		    grid.DataSource = table
  		    HoldFirstColumn(grid)
  		    SetViewSize(grid, Me.userRecordManager.GetUserRecordColumnsInfo)
  		    SetColorToOnlyTotalRow(grid)
  		  End If
  		End If
    Catch ex As Exception
      MsgBox.ShowError(ex)
    End Try
	End Sub
	
	''' <summary>
	''' 現在表示されているページのDataGridViewを返す。
	''' </summary>
	Function GetShowingDataGridViewInPersonalDataPage() As DataGridView
	  Dim tabPage As TabPage = Me.tabInPersonalTab.SelectedTab
		Return TabPageUtils.GetDataGridView(tabPage)
	End Function
	
	''' <summary>
	''' 現在表示されているデータの名前を取得する
	''' </summary>
	''' <returns></returns>
	Function GetShowingDataNameInPersonalDataPage() As String
	  Dim userInfo As UserInfo = GetSelectedUserInfo()
	  Dim month As Integer = GetSelectedPageMonthInPersonalDataPage()
	  
    Dim pageName As String
	  If month > 0 Then
	    pageName = Me.excelProperties.SheetName(month)
	  Else
	    pageName = PAGE_NAME_TOTAL
	  End If
	  
	  Return userInfo.GetSimpleId & userInfo.GetName & "_" & pageName
	End Function
	
	''' <summary>
	''' 現在選択されているページの月を取得する。
	''' 月データのページでない場合は-1を返す。
	''' </summary>
	Private Function GetSelectedPageMonthInPersonalDataPage() As Integer
		Dim tabPage As TabPage = Me.tabInPersonalTab.SelectedTab
		
		' 現在ページの月を取得する
		Dim currentTerm As DateTerm = Nothing
		Dim month As Integer = -1
		If Me.pageNameAndTermDictionary.TryGetValue(tabPage.Text, currentTerm) Then
		  month = currentTerm.BeginDate.Month
		End If
		
		Return month
	End Function
	
	''' <summary>
	''' 現在選択されているユーザ情報を取得する。
	''' </summary>
	Private Function GetSelectedUserInfo() As UserInfo
	  Return DirectCast(Me.cboxUserName.SelectedValue, UserInfo)
	End Function
	
End Class
