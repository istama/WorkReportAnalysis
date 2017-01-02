'
' 日付: 2016/10/18
'
Imports System.Data
Imports System.Linq
Imports System.Collections.Concurrent
Imports Common.Account
Imports Common.Util
Imports Common.IO
Imports Common.COM
Imports Common.Extensions
Imports WorkReportAnalysis

''' <summary>
''' ユーザのレコードを格納する。
''' </summary>
Public NotInheritable Class UserRecord
  ''' ユーザのIDの数値
  Private ReadOnly idNumber As String
  ''' ユーザ名
  Private ReadOnly name As String
  
  ''' このレコードの列情報
  Private ReadOnly columnsInfo As UserRecordColumnsInfo
  ''' このクラスが保持するのデータの期間
  Private ReadOnly dateTerm As DateTerm
  
  ''' 各月ごとに分割されたデータを持つオブジェクト
  Private ReadOnly record As New ConcurrentDictionary(Of Integer, DataTable)
  
  Public Sub New(userinfo As UserInfo, recordColumnsInfo As UserRecordColumnsInfo, dataDateTerm As DateTerm)
    If userinfo   Is Nothing Then Throw New ArgumentNullException("userinfo is null")
    
    Dim begin As DateTime = dataDateTerm.BeginDate
    Dim _end  As DateTime = dataDateTerm.EndDate
    Dim year  As Integer  = _end.Year - begin.Year
    If year > 1 OrElse (year = 1 AndAlso begin.Month <= _end.Month) Then
      Throw New ArgumentException(
        "データ期間は１年以内にしてください。" & vbCrLf &
        "BeginDate: " & begin.ToString("yyyy/MM/dd") & " EndDate: " & _end.ToString("yyyy/MM/dd"))
    End If
    
    
    Me.idNumber       = userinfo.GetSimpleId
    Me.name           = userinfo.GetName
    Me.columnsInfo    = recordColumnsInfo
    Me.dateTerm       = dataDateTerm
    Me.record         = CreateDataTables(Me.dateTerm)
  End Sub
  
  ''' <summary>
  ''' 指定した期間内の空のレコードを月単位で作成する。
  ''' </summary>
  Private Function CreateDataTables(dateTerm As DateTerm) As ConcurrentDictionary(Of Integer, DataTable)
    Dim dict As New ConcurrentDictionary(Of Integer, DataTable)
    
    ' 月単位でテーブルを作成する
    For Each m As DateTerm In dateTerm.MonthlyTerms
      Dim table As DataTable = Me.columnsInfo.CreateDataTable()
      Dim d As DateTime = m.BeginDate
      
      ' 日付の数だけテーブルの行を作成する
      For Each day As Integer In Enumerable.Range(1, DateTime.DaysInMonth(d.Year, d.Month))
        table.Rows.Add(table.NewRow)
      Next
      
      ' 月とテーブルをセットにして格納する
      dict.TryAdd(d.Month, table)
    Next
    
    Return dict
  End Function
  
  Public Function GetIdNumber() As String
    Return idNumber
  End Function
  
  Public Function GetName() As String
    Return name
  End Function
  
  Public Function GetExcelCulumnNodeTree As ExcelColumnNode
    Return Me.columnsInfo.CreateExcelColumnNodeTree
  End Function
  
  ''' <summary>
  ''' データの期間を取得する。
  ''' </summary>
  Public Function GetRecordDateTerm() As DateTerm
    Return Me.dateTerm 
  End Function
  
  ''' <summary>
  ''' 指定した月のテーブルを作成し返す。
  ''' </summary>
  Public Function GetRecord(month As Integer) As DataTable
    Dim table As DataTable = Nothing
    If Not Me.record.TryGetValue(month, table) Then
      Throw New ArgumentException("指定した月はレコードの範囲外です。 / month: " & month.ToString)
    End If
    
    Return table
  End Function
  
  ''' <summary>
  ''' 指定した月の１日単位のデータを取得する。
  ''' データ期間内に収まらない日のデータは含まれない。
  ''' </summary>
  Public Function GetDailyDataTable(year As Integer, month As Integer) As DataTable
    Dim min As New DateTime(year, month, 1)
    Dim max As New DateTime(year, month, DateTime.DaysInMonth(year, month))
    
    Return GetDailyDataTable(New DateTerm(min, max))
  End Function
  
  ''' <summary>
  ''' 指定した期間の１日単位のデータを取得する。
  ''' データ期間内に収まらない日のデータは含まれない。
  ''' </summary>
  Public Function GetDailyDataTable(dateTerm As DateTerm) As DataTable
    ' 新しいテーブルを作成する    
    Dim newTable As DataTable = Me.columnsInfo.CreateDataTable()
    ' 日付の範囲がこのデータの日付の範囲を越えていた場合、範囲内に収めるよう調整する
    Dim term As DateTerm = ModifyDateTerm(dateTerm)
    
    ' 各月のデータを新しいテーブルに書き込む
    For Each m As DateTerm In term.MonthlyTerms
      GetRecord(m.BeginDate.Month).
        Take(m.EndDate.Day).
        Skip(m.BeginDate.Day - 1).WriteTo(newTable)
    Next
    
    Return newTable
  End Function
  
  ''' <summary>
  ''' 指定した期間の１週間単位のデータを取得する。
  ''' </summary>
  Public Function GetWeeklyDataTable(dateTerm As DateTerm, exceptsRowUnfilled As Boolean) As DataTable
    ' 新しいテーブルを作成する
    Dim newTable As DataTable = Me.columnsInfo.CreateDataTable()
    
    ' 各週の合計値を新しいテーブルにセットする
    For Each w As DateTerm In ModifyDateTerm(dateTerm).WeeklyTerms()
      Dim newRow As DataRow = newTable.NewRow
      GetSumDataRow(w, exceptsRowUnfilled).CopyTo(newRow)
      newTable.Rows.Add(newRow)
    Next
    
    Return newTable
  End Function
  
  ''' <summary>
  ''' 指定した期間の１ヶ月単位のデータを取得する。
  ''' </summary>
  Public Function GetMonthlyDataTable(dateTerm As DateTerm, exceptsRowUnfilled As Boolean) As DataTable
    ' 新しいテーブルを作成する
    Dim newTable As DataTable = Me.columnsInfo.CreateDataTable()
    
    ' 各月の合計値を新しいテーブルにセットする
    For Each m As DateTerm In ModifyDateTerm(dateTerm).MonthlyTerms()
      Dim newRow As DataRow = newTable.NewRow
      GetSumDataRow(m, exceptsRowUnfilled).CopyTo(newRow)        
      newTable.Rows.Add(newRow)
    Next
    
    Return newTable
  End Function
  
  ''' <summary>
  ''' 指定した期間のデータの各列の合計値をセットした行データを返す。
  ''' 作業時間が０の作業件数を合計値から外すこともできる。
  ''' </summary>
  Public Function GetSumDataRow(dateTerm As DateTerm, exceptsRowUnfilled As Boolean) As DataRow
    ' 日付の範囲がこのデータの日付の範囲を越えていた場合、範囲内に収めるよう調整する
    Dim term As DateTerm = ModifyDateTerm(dateTerm)
    ' 指定した期間のデータをすべて取得する
    Dim table As DataTable = GetDailyDataTable(term)
    
    ' 作業時間が０の作業件数を合計値に含めるかどうか
    If exceptsRowUnfilled Then
      Return UserRecordCalculation.SumRowExceptedUnfilled(table, Me.columnsInfo)    
    Else
      Return table.SumRow()  
    End If
  End Function
  
  ''' <summary>
  ''' 日付の範囲がこのレコードの期間の範囲外だった場合、その範囲内におさめて返す。
  ''' </summary>
  Private Function ModifyDateTerm(term As DateTerm) As DateTerm
    If term.BeginDate > Me.dateTerm.EndDate OrElse term.EndDate < Me.dateTerm.BeginDate Then
      Throw New ArgumentException("指定した期間がこのレコードの期間の範囲外です。 / term: " & term.ToString)
    End If
    
    Dim beginDate As DateTime = term.BeginDate
    If beginDate < Me.dateTerm.BeginDate Then
      beginDate = Me.dateTerm.BeginDate
    End If
    
    Dim endDate As DateTime = term.EndDate
    If endDate > Me.dateTerm.EndDate Then
      endDate = Me.dateTerm.EndDate
    End If
    
    Return New DateTerm(beginDate, endDate)
  End Function
  
End Class
