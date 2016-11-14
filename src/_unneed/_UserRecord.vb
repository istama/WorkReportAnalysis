''
'' 日付: 2016/06/22
''
'Option Strict Off
'
'Imports System.Data
'Imports System.Linq
'Imports System.Threading
'Imports Common.Util
'Imports Common.Threading
'Imports Common.Data
'
'Public Class UserRecord
'	Private Shared COLUMN_NAME_COUNT As String        = "件数"
'	Private Shared COLUMN_NAME_WORKTIME As String     = "作業時間"
'	Private Shared COLUMN_NAME_PRODUCTIVITY As String = "生産性"
'	Private Shared COLUMN_NAME_WORKDAY As String      = "出勤日"
'	
'	Private properties As MyProperties	
'	Private dataSet As DataSet
'	
'	Private columns As DataColumnCollection
'	
'	Public Sub New(properties As MyProperties)
'		If properties Is Nothing Then Throw New ArgumentNullException("properties is null")
'		
'		Me.properties = properties
'		Me.dataSet    = New DataSet
'		
'		Me.columns    = Me.CreateTable().Columns
'	End Sub
'	
'	''' <summary>
'	''' テーブルを作成する。
'	''' </summary>
'	Public Function CreateTable() As DataTable
'		Dim columnNode As ColumnNode = UserRecord.CreateColumnNode(properties)
'		
'		Dim table As DataTable       = ColumnNode.CreateDataTable(columnNode.Children)
'		' ノードのルートである出勤日フラグはテーブルの１番後ろの列にする
'		Dim newcol As DataColumn     = ColumnNode.CreateColumn(columnNode)
'		table.Columns.Add(newcol)
'		
'		Return table
'	End Function
'	
'	''' <summary>
'	''' テーブルを作成する。
'	''' 指定した月がテーブルの名前にセットされ、テーブルはこのインスタンスに紐付けられる。
'	''' </summary>
'	Public Function CreateTable(month As Integer) As DataTable
'		If month < 1 OrElse month > 12 Then Throw New ArgumentException("the month is invalid. / " & month)
'		
'		Dim tableName As String = Me.properties.GetSheetName(month.ToString)
'		
'		If dataSet.Tables.Contains(tableName) Then
'			Throw New ArgumentException("table of the month has already exists.")
'		End If
'		
'		Dim table As DataTable = Me.CreateTable()
'		table.TableName = tableName
'		
'		Me.dataSet.Tables.Add(table)
'		
'		Return table		
'	End Function
'	
'	''' <summary>
'	''' 列名のコレクションを取得する。
'	''' </summary>
'	Public Function TableColumn() As DataColumnCollection
'		Return Me.columns
'	End Function
'	
'	''' <summary>
'	''' 指定した名前のシートを持っているか判定する。
'	''' </summary>
'	Public Function HasTable(tableName As String) As Boolean
'		If tableName Is Nothing Then Throw New ArgumentNullException("tableName is null")
'		
'		Return dataSet.Tables(tableName) IsNot Nothing
'	End Function
'	
'	''' <summary>
'	''' 指定した月のテーブルを取得する。
'	''' </summary>
'	Public Function GetTable(month As Integer) As DataTable
'		Return GetTable(properties.GetSheetName(month.ToString))
'	End Function
'	
'	''' <summary>
'	''' 指定した名前のテーブルを取得する。
'	''' </summary>
'	Public Function GetTable(tableName As String) As DataTable
'		If tableName Is Nothing Then Throw New ArgumentNullException("tableName is null")
'		
'		Dim table As DataTable = dataSet.Tables(tableName)
'		If table Is Nothing Then
'			Throw New KeyNotFoundException("the tableName dose not exists in the DataSet / " & tableName)
'		End If
'		
'		Return table.Copy
'	End Function
'	
'	''' <summary>
'	''' 指定した期間内のデータを抜き出す。
'	''' </summary>
'	Public Function GetTable(term As DateTerm) As DataTable
'		Return GetTable(term.Min, term.Max)
'	End Function
'	
'	Public Function GetTable(begin As DateTime, _end As DateTime) As DataTable
'		If begin > _end Then Throw New ArgumentException("begin table is later than _end")
'		
'		Dim newTable As DataTable = CreateTable()
'		
'		For Each d In DateUtils.GetDateListOfEveryMonth(begin, _end)
'			Dim tableName As String = Me.properties.GetSheetName(d.Month.ToString)
'			Dim table As DataTable = Me.dataSet.Tables(tableName)
'			
'			' 現在の月でデータを読み込む開始行をセットし、データを読み込む行数をセットする
'			Dim startRow As Integer
'			Dim allRowCount As Integer
'			If d.Year = begin.Year AndAlso d.Month = begin.Month Then
'				startRow = begin.Day - 1
'				If d.Month = _end.Month Then
'					allRowCount = _end.Day - startRow
'				Else
'					allRowCount = DateTime.DaysInMonth(begin.Year, begin.Month) - startRow
'				End If
'			ElseIf d.Year = _end.Year AndAlso d.Month = _end.Month
'				startRow = 0
'				allRowCount = _end.Day
'			Else
'				startRow = 0
'				allRowCount = DateTime.DaysInMonth(d.Year, d.Month)
'			End If
'			
'			For Each r As DataRow In (From n In table Select n).Skip(startRow).Take(allRowCount)
'				newTable.ImportRow(r)
'			Next
'		Next
'		
'		Return newTable
'	End Function
'	
'	''' <summary>
'	''' 指定した日付の行をresultTableに追加する。
'	''' </summary>
'	Public Sub AddRowToTable(d As DateTime, resultTable As DataTable)
'		Dim table As DataTable = Me.dataSet.Tables(properties.GetSheetName(d.Month.ToString))
'		Dim row As DataRow = table.Rows(d.Day - 1)
'		SyncLock resultTable
'			resultTable.ImportRow(row)
'		End SyncLock
'	End Sub
'	
'	''' <summary>
'	''' 指定した期間中のデータの各列の合計を求め、それを行にしてresultTableに追加する。
'	''' isDependingをTrueにすると、作業時間列に値がない行の件数列の値は、合計値に含めない。
'	''' </summary>
'	Public Sub AddTotalRowToTable(term As DateTerm, isDepending As Boolean, resultTable As DataTable)
'		' 開始日と終了日が同じ場合
'		If term.Min = term.Max Then
'			AddRowToTable(term.Min, resultTable)
'		Else
'			AddTotalRowToTable(GetTable(term), isDepending, resultTable)
'		End If
'	End Sub
'	
'	''' <summary>
'	''' 指定したテーブルの各列の合計を求め、それを行にしてテーブルに追加する。
'	''' isDependingをTrueにすると、作業時間列に値がない行の件数列の値は、合計値に含めない。
'	''' </summary>
'	Public Sub AddTotalRowToTable(table As DataTable, isDepending As Boolean)
'		AddTotalRowToTable(table, isDepending, table)
'	End Sub
'	
'	''' <summary>
'	''' 指定したテーブルの各列の合計を求め、それを行にしてresultTableに追加する。
'	''' isDependingをTrueにすると、作業時間列に値がない行の件数列の値は、合計値に含めない。
'	''' </summary>
'	Public Sub AddTotalRowToTable(table As DataTable, isDepending As Boolean, resultTable As DataTable)
''		Dim columns As DataColumnCollection = table.Columns
''		Dim resultRow As DataRow = resultTable.NewRow
''		
''		For i = 0 To columns.Count - 1
''			Dim txt As String = String.Empty
''			Dim colName As String = columns(i).ToString
''			If colName.Contains(COLUMN_NAME_COUNT) Then
''				If isDepending Then
''					Dim total As Double = TotalColumnValues(table, colName, columns(i+1).ToString)
''					txt = Math.Round(total).ToString
''				Else
''					Dim total	As Double = TotalColumnValues(table, colName, Nothing)
''					txt = Math.Round(total).ToString					
''				End If
''			ElseIf colName.Contains(COLUMN_NAME_WORKTIME)
''				Dim total As Double = TotalColumnValues(table, colName, Nothing)
''				txt = Math.Round(total, 2, MidpointRounding.AwayFromZero).ToString
''			ElseIf colName.Contains(COLUMN_NAME_PRODUCTIVITY)
''				Dim s1 As String = DirectCast(resultRow(columns(i - 2).ToString), String)
''				Dim s2 As String = DirectCast(resultRow(columns(i - 1).ToString), String)
''				Dim v1 As Double
''				Dim v2 As Double
''				If Double.TryParse(s1, v1) AndAlso Double.TryParse(s2, v2) AndAlso v2 <> 0.0 Then
''					Dim ave As Double = Math.Round(v1 / v2, 2, MidpointRounding.AwayFromZero) 
''					txt = ave.ToString
''				End If
''			End If
''			
''			resultRow(columns(i).ToString) = txt
''		Next
''		
''		SyncLock resultTable
''			resultTable.Rows.Add(resultRow)		
''		End SyncLock
'
'		SyncLock resultTable
'			Dim row As DataRow = resultTable.NewRow
'			TotalRow(table, isDepending, row)
'			resultTable.Rows.Add(row)		
'		End SyncLock
'	End Sub	
'	
'	Public Sub TotalRow(table As DataTable, isDepending As Boolean, resultRow As DataRow)
'		Dim columns As DataColumnCollection = table.Columns
'		
'		For i = 0 To columns.Count - 1
'			Dim txt As String = String.Empty
'			Dim colName As String = columns(i).ToString
'			
'			If colName.Contains(COLUMN_NAME_COUNT) Then
'				If isDepending Then
'					Dim total As Double = TotalColumnValues(table, colName, columns(i+1).ToString)
'					txt = Math.Round(total).ToString
'				Else
'					Dim total	As Double = TotalColumnValues(table, colName, Nothing)
'					txt = Math.Round(total).ToString					
'				End If
'			ElseIf colName.Contains(COLUMN_NAME_WORKTIME)
'				Dim total As Double = TotalColumnValues(table, colName, Nothing)
'				txt = Math.Round(total, 2, MidpointRounding.AwayFromZero).ToString
'			ElseIf colName.Contains(COLUMN_NAME_PRODUCTIVITY)
'				Dim s1 As String = resultRow(columns(i - 2).ToString).ToString
'				Dim s2 As String = resultRow(columns(i - 1).ToString).ToString
'				Dim v1 As Double
'				Dim v2 As Double
'				If Double.TryParse(s1, v1) AndAlso Double.TryParse(s2, v2) AndAlso v2 <> 0.0 Then
'					Dim ave As Double = Math.Round(v1 / v2, 2, MidpointRounding.AwayFromZero) 
'					txt = ave.ToString
'				End If
'			End If
'			
'			Select Case table.Columns(columns(i).ToString).DataType
'				Case GetType(Double)
'					Dim v As Double
'					If Double.TryParse(txt, v) AndAlso v <> 0 Then
'						resultRow(columns(i).ToString) = v
'					End If				
'				Case GetType(Integer)
'					Dim v As Integer
'					If Integer.TryParse(txt, v) AndAlso v <> 0  Then
'						resultRow(columns(i).ToString) = v
'					End If				
'				Case Else
'					resultRow(columns(i).ToString) = txt
'			End Select			
'		Next
'		
'	End Sub
'	
'	''' <summary>
'	''' 指定したテーブルの指定した列の合計を求める。
'	''' dependingColumnNameに列名を指定した場合、その列に値がない行の値は合計値に含めない。
'	''' </summary>
'	Public Function TotalColumnValues(table As DataTable, columnName As String, dependingColumnName As String) As Double
'		If table      Is Nothing Then Throw New ArgumentNullException("table is null")
'		If columnName Is Nothing Then Throw New ArgumentNullException("columnName is null")
'		
'		Dim rows As IEnumerable(Of DataRow) = New EnumerableList(Of DataRow)(table.Rows)
'		Dim filter As Func(Of DataRow, Boolean)
'		If dependingColumnName Is Nothing Then
'			filter = Function(row) True
'		Else
'			filter =
'				Function(row)
'					Dim d As String = row(dependingColumnName).ToString
'					Dim dCnt As Integer
'					Return Not String.IsNullOrEmpty(d) AndAlso Double.TryParse(d, dCnt) AndAlso dCnt <> 0.0
'				End FUnction
'		End If
'		
'		Dim total As Integer = 0
'	
'		MultiTask.Run(Of DataRow)(
'			rows,
'			filter,	
'			Function(row)
'				Return _
'					Sub(o)
'						Dim cnt As Double
'						If Double.TryParse(row(columnName).ToString, cnt) Then
'							cnt *= 1000
'							Interlocked.Add(total, CType(cnt, Integer))
'						End If
'					End Sub
'			End Function,
'			Nothing
'		)
'		
''		For Each row In table.Rows
''			' 依存する列がある場合
''			If dependingColumnName IsNot Nothing Then
''				' 依存する列の値が空や0の場合、この行の値は読み込まない
''				Dim d As String = row(dependingColumnName).ToString
''				Dim dCnt As Integer
''				If String.IsNullOrEmpty(d) OrElse Not Double.TryParse(d, dCnt) OrElse dCnt = 0.0 Then	
''					Continue For
''				End If
''			End If
''			
''			Dim cnt As Double
''			If Double.TryParse(row(columnName).ToString, cnt) Then
''				total += cnt
''			End If
''		Next
'		
'		Return total / 1000
'	End Function
'	
'	''' <summary>
'	''' 指定したテーブルの指定した列の平均を求める。
'	''' </summary>
'	Public Function AverageColumnValues(table As DataTable, columnName As String) As Double
'		If table      Is Nothing Then Throw New ArgumentNullException("table is null")
'		If columnName Is Nothing Then Throw New ArgumentNullException("columnName is null")
'		
'		Dim ave As Double = 0.0
'		Dim validRowsCnt As Integer = 0
'		
'		For Each row In table.Rows
'			Dim cnt As Double = 0.0
'			Dim txt As String = row(columnName).ToString
'			If Double.TryParse(txt, cnt) Then
'				ave += cnt
'				validRowsCnt += 1
'			End If
'		Next
'		
'		Return ave / validRowsCnt
'	End Function	
'	
'	''' <summary>
'	''' プロパティファイルから列ノードを作成。
'	''' </summary>
'	Public Shared Function CreateColumnNode(prop As MyProperties) As ColumnNode
'		' 出勤日フラグの列ノードを作成
'		Dim workdayNode As New ColumnNode(COLUMN_NAME_WORKDAY, prop.GetWorkDayCol, GetType(String))
'		
'		' 各作業項目の列ノードを作成し、出勤日列ノードの子に加える
'		For Each v As MyProperties.ItemValues In prop.GetItemValuesList
'			If Not String.IsNullOrEmpty(v.Name) AndAlso _
'				 Not String.IsNullOrEmpty(v.WorkCountCol) AndAlso _
'				 Not String.IsNullOrEmpty(v.WorkTimeCol) AndAlso _
'				 Not String.IsNullOrEmpty(v.ProductivityCol) Then
'				workdayNode.
'					AddChild(v.Name & vbCrLf & COLUMN_NAME_COUNT,        v.WorkCountCol,    GetType(Integer)).
'					AddChild(v.Name & vbCrLf & COLUMN_NAME_WORKTIME,     v.WorkTimeCol,     GetType(Double)).
'					AddChild(v.Name & vbCrLf & COLUMN_NAME_PRODUCTIVITY, v.ProductivityCol, GetType(Double))				
'			End If
'		Next
'		
'		' 備考の列ノード作成し、出勤日フラグの列ノードの子に追える
'		If Not String.IsNullOrEmpty(prop.GetNoteName) AndAlso Not String.IsNullOrEmpty(prop.GetNoteCol) Then
'			workdayNode.AddChild(prop.GetNoteName, prop.GetNoteCol, GetType(String))
'		End If
'		
'		Return workdayNode
'	End Function
'	
'
'End Class
