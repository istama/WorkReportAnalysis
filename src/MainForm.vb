'
' 日付: 2016/06/18
'
Imports System.Data
Imports Common.account
Imports Common.Util
Imports Common.Extensions
Imports Common.COM
Imports Common.IO

Public Partial Class MainForm
  Private Const PROPERTY_FILE_PATH       = ".\application.properties"
  Private Const EXCEL_PROPERTY_FILE_PATH = ".\excel.properties"
	
	Private Const TABPAGE_NAME_PERSONAL = "個人データ"
	Private Const TABPAGE_NAME_DATE     = "日付データ"
	Private Const TABPAGE_NAME_TOTAL    = "集計データ"
	
	Private initialized As Boolean
	
	' アプリケーションのプロパティファイル
	Private properties As New MyProperties(PROPERTY_FILE_PATH)
	' Excelのプロパティファイル
	Private excelProperties As New ExcelProperties(EXCEL_PROPERTY_FILE_PATH)
	
	' ユーザ情報管理オブジェクト
	Private userInfoManager As UserInfoManager
	Private userRecordManager As UserRecordManager
	
  ' このアプリケーションで取り扱うデータの期間
	Private dateTerm As DateTerm
	
	' ファイル保存ダイアログ
	Private saveFileDialog As New MySaveFileDialog
	
	Public Sub New()
		Me.InitializeComponent()
		Me.initialized = False
	End Sub
	
	Sub MainFormLoad(sender As Object, e As EventArgs)
	  Try
	    Log.SetFilePath(".\log.txt")
	    
	    ' ユーザ情報を管理するオブジェクトを作成
	    Me.userInfoManager = UserInfoManager.Create(Me.properties.UserFilePath)
	    Me.userRecordManager = New UserRecordManager(Me.excelProperties)
	    
	    LoadExcel()
	    
			dateTerm = New DateTerm(Me.excelProperties.BeginDate, Me.excelProperties.EndDate)

      ' 各コンボボックスの要素を初期化
			InitTabRoot()
			InitCboxUserNames()
			InitPersonalDataGridView()
			InitDateDataGridView()
			InitTallyDataGridView()

			' ウィンドウを最大化する
			Me.WindowState = FormWindowState.Maximized
			
			Me.initialized = True
			
			' アプリケーション画面の表示時に一番最初に表示されるユーザレコード
			CboxUserNameSelectedIndexChanged(Me.cboxUserName, New EventArgs)
		Catch ex As Exception
			MsgBox.ShowError(ex)
		End Try
	End Sub
	
	Private Sub LoadExcel()
	  Dim res As DialogResult = 
	    MessageBox.Show(
	      "Excelファイルを読み込みます。" & vbCrLf & "この処理は時間がかかることがありますがよろしいですか？",
	      "確認",
	      MessageBoxButtons.OKCancel,
	      MessageBoxIcon.Information)
	    
	  If res = DialogResult.OK Then
  	  ProgressBarForm.Loader = New Loader(Me.userRecordManager, Me.userInfoManager, Me.excelProperties)
  	  ProgressBarForm.ShowDialog()
  	End If
	End Sub
	
	Private Sub InitTabRoot()
		Me.tabRoot.TabPages(0).Text = TABPAGE_NAME_PERSONAL
		Me.tabRoot.TabPages(1).Text = TABPAGE_NAME_DATE
		Me.tabRoot.TabPages(2).Text = TABPAGE_NAME_TOTAL
	End Sub
	
	''' <summary>
	''' ユーザ名のコンボボックスの要素を初期化する。
	''' </summary>
  Private Sub InitCBoxUserNames()
    InitComboBox(
      Me.cboxUserName,
      Me.userInfoManager.UserInfos,
      GetType(UserInfo),
      Function(ui) ui.GetSimpleId.ToString() & " " & ui.GetName.ToString()) ' コンボボックスの要素として表示される文字列
	End Sub
	
	''' <summary>
	''' コンボボックスの要素を初期化する。
	''' </summary>
	''' <param name="cbox">初期化するコンボボックス</param>
	''' <param name="values">コンボボックスの要素のコレクション</param>
	''' <param name="typeOfValue">要素の型</param>
	''' <param name="f">コンボボックスに表示される文字列を返す関数</param>
	Private Sub InitComboBox(Of T)(cbox As ComboBox, values As IEnumerable(Of T), typeOfValue As Type, f As Func(Of T, String))
	  If typeOfValue Is Nothing Then Throw New ArgumentNullException("typeOfValue is null")
	  If f Is Nothing Then Throw New ArgumentNullException("f is null")
	  
	  If cbox IsNot Nothing AndAlso values IsNot Nothing Then
			Dim DISPLAY As String = "d"
			Dim VALUE   As String = "v"
			
			Dim dataTable As New DataTable
			dataTable.Columns.Add(DISPLAY, GetType(String))
			dataTable.Columns.Add(VALUE,   typeOfValue)
			
			values.ForEach(
			  Sub(e)
					Dim row As DataRow = dataTable.NewRow
					row(DISPLAY) = f(e) ' コンボボックスの要素として表示される文字列
					row(VALUE)   = e	' 選択した要素から取得できる値
					dataTable.Rows.Add(row)			  
			  End Sub)
			
			dataTable.AcceptChanges
			cbox.DataSource    = dataTable
			cbox.DisplayMember = DISPLAY
			cbox.ValueMember   = VALUE	    
    End If
	End Sub
	
	''' <summary>
	''' １番外側のタブページが選択された場合に発生するイベント。
	''' </summary>
	Sub TabRoot_SelectedIndexChanged(sender As Object, e As EventArgs)
		If initialized = False Then Return
		' 現在開かれているページのグリッドビューを表示する
		ShowGridView()
	End Sub
	
	''' <summary>
	''' 現在開かれているページのグリッドビューを表示する。
	''' </summary>
	Private Sub ShowGridView()
		Dim pageName = Me.tabRoot.SelectedTab.Text
		If pageName = TABPAGE_NAME_PERSONAL Then
			ShowPersonalDataGridView()
		ElseIf pageName = TABPAGE_NAME_DATE
			ShowDateDataGridView()
		ElseIf pageName = TABPAGE_NAME_TOTAL
			ShowTallyDataGridView()
		End If		
	End Sub
	
	''' <summary>
	''' 個人ページのユーザ名のコンボボックスが選択された場合に発生するイベント。
	''' </summary>
	Sub CboxUserNameSelectedIndexChanged(sender As Object, e As EventArgs)
		If initialized = False Then	Return
		ShowPersonalDataGridView()
	End Sub
	
	''' <summary>
	''' 日付ページのカレンダーが選択された場合に発生するイベント。
	''' </summary>
	Sub DateTimePickerInDatePageValueChanged(sender As Object, e As EventArgs)
		If initialized = False Then Return
		ShowDateDataGridView()
	End Sub
	
	''' <summary>
	''' 日付ページの週コンボボックスが選択された場合に発生するイベント。
	''' </summary>
	Sub CboxWeeklySelectedIndexChanged(sender As Object, e As EventArgs)
		If initialized = False Then Return
		ShowDateDataGridView()
	End Sub
	
	''' <summary>
	''' 日付ページの月コンボボックスが選択された場合に発生するイベント。
	''' </summary>
	Sub CboxMonthlySelectedIndexChanged(sender As Object, e As EventArgs)
		If initialized = False Then Return
		ShowDateDataGridView()		
	End Sub
	
	''' <summary>
	''' 集計ページの月コンボボックスが選択された場合に発生するイベント。
	''' </summary>
	Sub CboxTallyMonthlySelectedIndexChanged(sender As Object, e As EventArgs)
		If initialized = False Then Return
		ShowTallyDataGridView()
	End Sub
	
	''' <summary>
	''' 作業時間が入力されていない行の件数を合計値に含めるかどうかのチェックボックスが選択された場合に発生するイベント。
	''' </summary>
	Sub ChkBoxExcludeDataCheckedChanged(sender As Object, e As EventArgs)
		If initialized = False Then Return
		ShowGridView()
	End Sub
	
  ''' <summary>
  ''' CSV出力ボタンがクリックされた場合に発生するイベント
  ''' </summary>
	Sub BtnOutputCsvClick(sender As Object, e As EventArgs)
	  If initialized = False Then Return
	  
	  ' 現在表示されているDataGridViewを取得する
	  Dim grid As DataGridView = GetShowingDataGridView()
	  If grid IsNot Nothing Then
	    Dim table As DataTable = DirectCast(grid.DataSource, DataTable)
	    
	    ' 現在表示されているデータの名前を取得する
	    Dim fileName As String         = GetShowingDataName() & ".csv"
	    ' 保存ダイアログを開き、OKが押されたら出力ストリームを取得する
	    Dim stream   As SaveFileStream = Me.saveFileDialog.Save(fileName)
	    If stream IsNot Nothing Then
	      Try
	        ' DataTableをCSVに変換してファイルに出力する
	        stream.Open()
	        table.ToCSV().ForEach(Sub(csv) stream.Write(csv))
	      Finally
	        stream.Close()
	      End Try
	    End If
  	End If
	End Sub
	
	''' <summary>
	''' 再読み込みボタンがクリックされた場合に発生するイベント。
	''' </summary>
	Sub BtnReloadClick(sender As Object, e As EventArgs)
    LoadExcel()	  
	End Sub
	
	''' <summary>
	''' 閉じるボタンがクリックされた場合に発生するイベント。
	''' </summary>
	Sub BtnCloseClick(sender As Object, e As EventArgs)
		' アプリケーションを閉じる
		Me.Close
	End Sub
	
  ''' <summary>
  ''' フォームを閉じる直前に呼び出される。
  ''' </summary>
  Sub MainFormClosing(sender As Object, e As System.ComponentModel.CancelEventArgs)
  End Sub

End Class

