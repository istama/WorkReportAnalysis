'
' 日付: 2016/10/18
'
Imports System.Data
Imports System.Linq
Imports System.Collections.Concurrent
Imports Common.Account
Imports Common.Util
Imports Common.IO
Imports Common.Extensions

''' <summary>
''' ユーザレコードを一元管理し、アプリケーション上に表示する表形式に加工して返す。
''' </summary>
Public Class UserRecordManager
  ''' Excelの設定  
  Private ReadOnly properties As ExcelProperties
  
  ''' テーブルの列の構造や情報をまとめたオブジェクト
  Private ReadOnly recordColumnsInfo As UserRecordColumnsInfo
  
  ''' ユーザレコードを保持するオブジェクト。
  Private ReadOnly userRecordBuffer As UserRecordBuffer
  ''' ユーザレコードを読み込むオブジェクト。
  Private ReadOnly userRecordLoader As UserRecordLoader
  
  Public Sub New(properties As ExcelProperties)
    If properties Is Nothing Then Throw New ArgumentNullException("properties is null")
    
    Me.properties        = properties
    Me.recordColumnsInfo = UserRecordColumnsInfo.Create(properties)
    Me.userRecordLoader  = New UserRecordLoader(properties)
    Me.userRecordBuffer  = Me.userRecordLoader.GetUserRecordBuffer()
  End Sub
  
  Public Function Loader() As UserRecordLoader
    Return Me.userRecordLoader
  End Function
  
  ''' <summary>
  ''' 指定したユーザが登録されているか判定する。
  ''' </summary>
  Public Function Stored(userInfo As UserInfo) As Boolean
    If userInfo Is Nothing Then Throw New ArgumentNullException("userInfo is null")
    
    Return Me.userRecordBuffer.Stored(userInfo)
  End Function
  
  Public Function GetUserRecordColumnsInfo() As UserRecordColumnsInfo
    Return Me.recordColumnsInfo
  End Function
  

  ''' <summary>
  ''' 指定したユーザの指定した月のレコードを取得する。
  ''' </summary>  
  Public Function GetDailyRecord(userInfo As UserInfo, year As Integer, month As Integer, exceptsRowUnfilled As Boolean) As DataTable
    If userInfo Is Nothing Then Throw New ArgumentNullException("userInfo is null")
    
    Dim min  As New DateTime(year, month, 1)
    Dim max  As New DateTime(year, month, DateTime.DaysInMonth(year, month))
    Dim term As New DateTerm(min, max)
    
    Dim record As UserRecord = Me.userRecordBuffer.GetUserRecord(userInfo)
    Dim table  As DataTable  = GetDailyRecordNotContainedSumRow(record, term)
    ' 合計行を追加する
    AddSumRow(table, exceptsRowUnfilled)
    
    Return table
  End Function
  
  ''' <summary>
  ''' 指定した期間の１日ごとの集計データを収めたテーブルを取得する。
  ''' 集計データとは、全ユーザのデータを合計したデータのこと。
  ''' </summary>
  Public Function GetDailyTotalRecord(term As DateTerm, exceptsRowUnfilled As Boolean) As DataTable
    ' 合計レコードを取得する
    Dim totalRecord As UserRecord = GetTotalRecord(exceptsRowUnfilled)
    Dim table       As DataTable  = GetDailyRecordNotContainedSumRow(totalRecord, term)
    ' 合計行を追加する
    AddSumRow(table, exceptsRowUnfilled)
    
    Return table
  End Function
  
  ''' <summary>
  ''' 渡したレコードの指定した月のテーブルを取得する。
  ''' レコードには合計行は含まれない。
  ''' </summary>
  Private Function GetDailyRecordNotContainedSumRow(record As UserRecord, term As DateTerm) As DataTable
    If record Is Nothing Then Throw New ArgumentNullException("record is null")
    
    Dim table         As DataTable  = record.GetDailyDataTable(term)
    Dim labelingTable As DataTable  = CreateDataTableLabelingDate(table, term.DailyTerms(Function(d) d.Day & "日"))
    
    ' 生産性を計算し列にセットする
    CalcProductivity(labelingTable)
    
    Return labelingTable
  End Function
  
  ''' <summary>
  ''' 指定したユーザの指定した期間のレコードを１週間単位で取得する。
  ''' </summary>
  Public Function GetWeeklyRecord(userInfo As UserInfo, dateTerm As DateTerm, exceptsRowUnfilled As Boolean) As DataTable
    If UserInfo Is Nothing Then Throw New ArgumentNullException("userInfo is null")
    
    Dim record As UserRecord = Me.userRecordBuffer.GetUserRecord(userInfo)
    Dim table  As DataTable  = GetWeeklyRecordNotContainedSumRow(record, dateTerm, exceptsRowUnfilled)
    ' 合計行を追加する
    AddSumRow(table, exceptsRowUnfilled)
    
    Return table
  End Function
  
  ''' <summary>
  ''' 指定した期間の集計レコードを１週間単位で取得する。
  ''' </summary>
  Public Function GetWeeklyTotalRecord(dateTerm As DateTerm, exceptsRowUnfilled As Boolean) As DataTable
    ' 合計レコードを取得する
    Dim record As UserRecord = GetTotalRecord(exceptsRowUnfilled)
    Dim table  As DataTable  = GetWeeklyRecordNotContainedSumRow(record, dateTerm, exceptsRowUnfilled)
    ' 合計行を追加する
    AddSumRow(table, exceptsRowUnfilled)
    
    Return table
  End Function
  
  ''' <summary>
  ''' 指定したユーザの指定した期間のレコードを１週間単位で取得する。
  ''' レコードには合計行は含まれない。
  ''' </summary>
  Private Function GetWeeklyRecordNotContainedSumRow(record As UserRecord, dateTerm As DateTerm, exceptsRowUnfilled As Boolean) As DataTable
    If record Is Nothing Then Throw New ArgumentNullException("record is null")
    
    Dim table         As DataTable  = record.GetWeeklyDataTable(dateTerm, exceptsRowUnfilled)
    Dim labelingTable As DataTable  = CreateDataTableLabelingDate(table, dateTerm.WeeklyTerms(DayOfWeek.Saturday, GetFuncForLabelingWeeklyTerms(dateTerm)))
    
    ' 生産性を計算し列にセットする
    CalcProductivity(labelingTable)
    
    Return labelingTable
  End Function
  
  ''' <summary>
  ''' DateTerm.WeeklyTerms()のラベルを生成する関数を返す。
  ''' </summary>
  Private Function GetFuncForLabelingWeeklyTerms(dateTerm As DateTerm) As Func(Of DateTime, DateTime, String)
    Dim weekCountInMonth = DateUtils.GetWeekCountInMonth(dateTerm.BeginDate, DayOfWeek.Saturday)
    Return _
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
      End Function      
  End Function
  
  ''' <summary>
  ''' 指定したユーザの指定した期間のレコードを１ヶ月単位で取得する。
  ''' </summary>
  Public Function GetMonthlyRecord(userInfo As UserInfo, dateTerm As DateTerm, exceptsRowUnfilled As Boolean) As DataTable
    If UserInfo Is Nothing Then Throw New ArgumentNullException("userInfo is null")
    
    Dim record As UserRecord = Me.userRecordBuffer.GetUserRecord(userInfo)
    Dim table  As DataTable  = GetMonthlyRecordNotContainedSumRow(record, dateTerm, exceptsRowUnfilled)
    ' 合計行を追加する
    AddSumRow(table, exceptsRowUnfilled)
    
    Return table
  End Function
  
  ''' <summary>
  ''' 指定した期間の集計レコードを月単位で取得する。
  ''' </summary>
  Public Function GetMonthlyTotalRecord(dateTerm As DateTerm, exceptsRowUnfilled As Boolean) As DataTable
    ' 合計レコードを取得する
    Dim record As UserRecord = GetTotalRecord(exceptsRowUnfilled)
    Dim table  As DataTable  = GetMonthlyRecordNotContainedSumRow(record, dateTerm, exceptsRowUnfilled)
    ' 合計行を追加する
    AddSumRow(table, exceptsRowUnfilled)
    
    Return table
  End Function  
  
  ''' <summary>
  ''' 指定したユーザの指定した期間のレコードを１ヶ月単位で取得する。
  ''' レコードには合計行は含まれない。
  ''' </summary>  
  Private Function GetMonthlyRecordNotContainedSumRow(record As UserRecord, dateTerm As DateTerm, exceptsRowUnfilled As Boolean) As DataTable
    If record Is Nothing Then Throw New ArgumentNullException("record is null")
    
    Dim table         As DataTable  = record.GetMonthlyDataTable(dateTerm, exceptsRowUnfilled)
    Dim labelingTable As DataTable  = CreateDataTableLabelingDate(table, dateTerm.MonthlyTerms(Function(b, e) b.Month & "月"))
    
    ' 生産性を計算し列にセットする
    CalcProductivity(labelingTable)
    
    Return labelingTable 
  End Function
  
  ''' <summary>
  ''' 指定したユーザの集計レコードを取得する。
  ''' </summary>
  Public Function GetSumRecord(userInfo As UserInfo, exceptsRowUnfilled As Boolean) As DataTable
    If userInfo Is Nothing Then Throw New ArgumentNullException("userInfo is null")
    
    Dim record As UserRecord = Me.userRecordBuffer.GetUserRecord(userInfo)
    
    Dim dateTerm        As DateTerm  = record.GetRecordDateTerm()
    Dim table           As DataTable = GetWeeklyRecordNotContainedSumRow(record, dateTerm, exceptsRowUnfilled)
    Dim monthlySumTable As DataTable = GetMonthlyRecordNotContainedSumRow(record, dateTerm, exceptsRowUnfilled)
    
    ' 合計行を追加する
    AddSumRow(monthlySumTable, exceptsRowUnfilled)
    
    ' tableの後ろの行にmonthlySumTableの行を追加する
    monthlySumTable.WriteTo(table)
    
    Return table
  End Function
  
  ''' <summary>
  ''' 各ユーザごとの指定した期間の集計レコードを集めたテーブルを取得する。
  ''' </summary>
  Public Function GetAllUserSumRecord(dateTerm As DateTerm, exceptsRowUnfilled As Boolean) As DataTable
    Dim allUserTable As DataTable = Me.GetUserRecordColumnsInfo.CreateDataTable(UserRecordColumnsInfo.NAME_COL_NAME)
    
    ' 各ユーザのレコードの集計値を計算し、新しいテーブルを作成する
    For Each record As UserRecord In Me.userRecordBuffer.GetUserRecordAll
      ' 指定した期間の値を集計した行を取得
      Dim sumRow As DataRow = record.GetSumDataRow(dateTerm, exceptsRowUnfilled)
      ' 集計テーブルの行に集計値をコピー      
      Dim newRow As DataRow = allUserTable.NewRow
      sumRow.CopyTo(newRow)
      ' ユーザ名の列にユーザ名をセット
      newRow(UserRecordColumnsInfo.NAME_COL_NAME) = record.GetIdNumber & " " & record.GetName
      
      allUserTable.Rows.Add(newRow)
    Next
    
    ' 生産性を計算し列にセットする
    CalcProductivity(allUserTable)
    
    ' 合計行を追加する
    AddSumRow(allUserTable, exceptsRowUnfilled)
    
    Return allUserTable
  End Function
  
  ''' <summary>
  ''' 集計レコードを取得する。
  ''' 作業時間が０の作業件数を集計に含めるかどうかを指定できる。
  ''' </summary>
  Private Function GetTotalRecord(exceptsRowUnfilled As Boolean) As UserRecord
    If exceptsRowUnfilled Then
      Return Me.userRecordBuffer.GetTotalRecordExceptedUnfilled()
    Else
      Return Me.userRecordBuffer.GetTotalRecord()
    End If    
  End Function
  
  ''' <summary>
  ''' 渡したテーブルの１列目に日付の列を追加して返す。
  ''' </summary>
  Private Function CreateDataTableLabelingDate(table As DataTable, terms As IEnumerable(Of DateTerm)) As DataTable
    Return CreateDataTableLabeling(table, UserRecordColumnsInfo.DATE_COL_NAME, terms.Select(Function(d) d.Label))
  End Function
  
  ''' <summary>
  ''' 渡したテーブルの１列目に名前の列を追加して返す。
  ''' </summary>
  Private Function CreateDataTableLabelingName(table As DataTable, names As IEnumerable(Of String)) As DataTable
    Return CreateDataTableLabeling(table, UserRecordColumnsInfo.NAME_COL_NAME, names)
  End Function  
  
  ''' <summary>
  ''' 渡したテーブルの１列目に指定した名前の列とその値をセットして返す。
  ''' </summary>
  Private Function CreateDataTableLabeling(table As DataTable, labelColName As String, labels As IEnumerable(Of String)) As DataTable
    Dim newTable As DataTable = Me.recordColumnsInfo.CreateDataTable(labelColName)
    
    ' 既存のテーブルのデータを新しいデータに書き込み
    table.WriteTo(newTable)
    ' 日付の列に日付をセットする
    Dim idx As Integer = 0
    For Each label As String In labels
      newTable.Rows(idx)(labelColName) = label
      idx += 1
      If idx >= newTable.Rows.Count Then
        Exit For
      End If
    Next
    
    Return newTable    
  End Function
  
  ''' <summary>
  ''' 各行の各作業項目の生産性を求めその列に追加する。
  ''' </summary>
  ''' <param name="table"></param>
  Private Sub CalcProductivity(table As DataTable)
    For Each dataRow As DataRow In table.AsEnumerable()
      CalcProductivity(dataRow)
    Next
  End Sub
  
  ''' <summary>
  ''' 各作業項目の生産性を求めその列に追加する。
  ''' </summary>
  Private Sub CalcProductivity(dataRow As DataRow)
    For Each cols As WorkItemColumnsInfo In Me.recordColumnsInfo.WorkItems
      Dim productColName As String = cols.WorkProductivityColInfo.name
      ' 生産性の列がある作業項目かどうか
      If Not String.IsNullOrWhiteSpace(productColName) Then
        Dim cntColName  As String = cols.WorkCountColInfo.name
        Dim timeColName As String = cols.WorkTimeColInfo.name
        ' 作業件数と作業時間の列に値が入っているかどうか
        If Not dataRow.IsNull(cntColName) AndAlso Not dataRow.IsNull(timeColName) Then
          Dim cnt  As Integer = dataRow.Field(Of Integer)(cntColName)
          Dim time As Double  = dataRow.Field(Of Double)(timeColName)
          ' 0以上なら計算する
          If cnt > 0 AndAlso time > 0.0 Then
            dataRow(productColName) = cnt / time
          End If
        End If
      End If
    Next
  End Sub
  
  ''' <summary>
  ''' テーブルにその合計行を追加する。
  ''' </summary>
  Private Sub AddSumRow(table As DataTable, exceptsRowUnfilled As Boolean)
    ' 合計行を求める
    Dim sumRow As DataRow = GetSumRow(table, exceptsRowUnfilled)
    
    ' 合計行に生産性の値をセットする
    CalcProductivity(sumRow)
    ' １列目にタイトルをセットする
    If sumRow.HasColumn(UserRecordColumnsInfo.DATE_COL_NAME) Then
      sumRow(UserRecordColumnsInfo.DATE_COL_NAME) = "合計"
    ElseIf sumRow.HasColumn(UserRecordColumnsInfo.NAME_COL_NAME)
      sumRow(UserRecordColumnsInfo.NAME_COL_NAME) = "合計"
    End If
    
    table.Rows.Add(sumRow)
  End Sub
  
  ''' <summary>
  ''' テーブルの各列の合計を求め、新しい行にセットして返す。
  ''' </summary>
  Private Function GetSumRow(table As DataTable, exceptsRowUnfilled As Boolean) As DataRow
    ' 作業時間が０の作業件数を合計値に含めるかどうか
    If exceptsRowUnfilled Then
      Return UserRecordCalculation.SumRowExceptedUnfilled(table, Me.recordColumnsInfo)    
    Else
      Return table.SumRow()  
    End If
  End Function
  
End Class
