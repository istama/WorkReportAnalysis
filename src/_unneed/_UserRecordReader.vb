''
'' 日付: 2016/06/20
''
'Imports System.Data
'Imports Common.COM
'Imports Common.account
'Imports Common.Util
'
'''' <summary>
'''' Excelファイルを読み込み、コレクションクラスに格納して返すクラス。
'''' </summary>
'Public Class UserRecordReader
'	Private reader As ExcelReader
'	Private properties As MyProperties
'	
'	''' 読み込む列をまとめたデータ構造
'	Private columnNode As ColumnNode
'	''' 読み込む月のリスト
'	Private readDateRange As List(Of DateTime)
'	
'	Public Sub New(properties As MyProperties)
'		Me.reader        = New ExcelReader(New Excel())
'		Me.properties    = properties
'		
'		Me.columnNode    = UserRecord.CreateColumnNode(properties)
'		Me.readDateRange = GetDateRange(properties)
'	End Sub
'	
'	Public Sub Init()
'		reader.Init
'	End Sub
'	
'	Public Sub Quit
'		reader.Quit		
'	End Sub
'	
'	''' <summary>
'	''' 指定したユーザのExcelファイルを読み込んでDataSetにして返す。
'	''' </summary>
'	Public Function Read(userID As String) As UserRecord
'		If userID Is Nothing Then
'			Throw New ArgumentException("userId is null")
'		End If
'		
'		' Excelファイルのパスを取得する
'		Dim filePath = properties.GetExcelFilePath(userID)
'		' 1日のデータの行
'		Dim firstDayRow As Integer = properties.GetRowOfFirstDayInAMonth
'		
'		'Dim dataSet As New DataSet()
'		Dim userRecord As New UserRecord(Me.properties)
'		
'		' 各月ごとにDataTabelを生成し、シートを読み込む
'		For Each d In readDateRange
'			' データテーブルを生成
'			' ノードのルートである出勤日フラグはテーブルの１番後ろの列にする
'			Dim table As DataTable = userRecord.CreateTable(d.Month)
'			
'			' １日ごとのデータを読み込む
'			For row As Integer = firstDayRow To (DateTime.DaysInMonth(d.Year, d.Month) - 1 + firstDayRow)
'				Dim dataRow As DataRow = table.NewRow()
'				reader.Read(filePath, table.TableName, row, ColumnNode, dataRow)
'				table.Rows.Add(dataRow)
'			Next
'		Next
'			
'		Return userRecord
'	End Function
'	
''	''' <summary>
''	''' プロパティファイルから列ノードを作成。
''	''' </summary>
''	Shared Function CreateColumnNode(prop As MyProperties) As ColumnNode
''		' 出勤日フラグの列ノードを作成
''		Dim workdayNode As New ColumnNode("出勤日", prop.GetWorkDayCol)
''		
''		' 各作業項目の列ノードを作成し、出勤日列ノードの子に加える
''		For Each v As MyProperties.ItemValues In prop.GetItemValuesList
''			If Not String.IsNullOrEmpty(v.Name) AndAlso _
''				 Not String.IsNullOrEmpty(v.WorkCountCol) AndAlso _
''				 Not String.IsNullOrEmpty(v.WorkTimeCol) AndAlso _
''				 Not String.IsNullOrEmpty(v.ProductivityCol) Then
''				workdayNode.
''					AddChild(v.Name & vbCrLf & "件数",     v.WorkCountCol).
''					AddChild(v.Name & vbCrLf & "作業時間", v.WorkTimeCol).
''					AddChild(v.Name & vbCrLf & "生産性",   v.ProductivityCol)				
''			End If
''		Next
''		
''		' 備考の列ノード作成し、出勤日フラグの列ノードの子に追える
''		If Not String.IsNullOrEmpty(prop.GetNoteName) AndAlso Not String.IsNullOrEmpty(prop.GetNoteCol) Then
''			workdayNode.AddChild(prop.GetNoteName, prop.GetNoteCol)
''		End If
''		
''		Return workdayNode
''	End Function
'	
'	''' <summary>
'	''' プロパティファイルに設定された開始日付と終了日付の間の日付オブジェクトを、
'	''' １ヶ月間隔で生成し、リストに格納して返す。
'	''' </summary>
'	Shared Function GetDateRange(prop As MyProperties) As List(Of DateTime)
'		Dim fromDate As DateTime = prop.GetSelectableMinDate
'		Dim toDate As DateTime   = prop.GetSelectableMaxDate
'		
'		Return DateUtils.GetDateListOfEveryMonth(fromDate, toDate)
'	End Function
'	
'End Class
