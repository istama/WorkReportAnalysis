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
  
  Public Sub New(userinfo As UserInfo, recordColumnsInfo As UserRecordColumnsInfo, properties As ExcelProperties)
    If userinfo   Is Nothing Then Throw New ArgumentNullException("userinfo is null")
    
    Me.idNumber       = userinfo.GetSimpleId
    Me.name           = userinfo.GetName
    Me.columnsInfo    = recordColumnsInfo
    Me.dateTerm       = New DateTerm(properties.BeginDate, properties.EndDate)
    Me.record         = CreateDataTables(Me.dateTerm)
  End Sub
  
  ''' <summary>
  ''' 指定した期間内の空のレコードを作成する。
  ''' </summary>
  Private Function CreateDataTables(dateTerm As DateTerm) As ConcurrentDictionary(Of Integer, DataTable)
    Dim dict As New ConcurrentDictionary(Of Integer, DataTable)
    
    dateTerm.MonthlyTerms.ForEach(
      Sub(m)
        Dim table As DataTable = Me.columnsInfo.CreateDataTable(Nothing)
        Dim d As DateTime = m.BeginDate
        Enumerable.Range(1, DateTime.DaysInMonth(d.Year, d.Month)).ForEach(
          Sub(day)
            Dim newRow As DataRow = table.NewRow
            table.Rows.Add(newRow)
        End Sub)
        
        dict.TryAdd(d.Month, table)
      End Sub)
    
    Return dict
  End Function
  
  Public Function GetIdNumber() As String
    Return idNumber
  End Function
  
  Public Function GetName() As String
    Return name
  End Function
  
  Public Function GetCulumnNodeTree As ExcelColumnNode
    Return Me.columnsInfo.CreateExcelColumnNodeTree
  End Function
  
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
  ''' </summary>
  Public Function GetDailyDataTableForAMonth(year As Integer, month As Integer) As DataTable
    Dim first As New DateTime(year, month, 1)
    Dim _end As New DateTime(year, month, Date.DaysInMonth(year, month))
    
    Return GetDailyDataTableLabelingDate(New DateTerm(first, _end), Function(term) term.BeginDate.Day & "日")
  End Function
  
  ''' <summary>
  ''' １列目に日付をつけて指定した期間の１日単位のデータを取得する。
  ''' </summary>
  Public Function GetDailyDataTableLabelingDate(dateTerm As DateTerm, f As Func(Of dateTerm, String)) As DataTable
    Dim table As DataTable = Me.columnsInfo.CreateDataTable(UserRecordColumnsInfo.DATE_COL_NAME)
    Dim term As DateTerm = ModifyDateTerm(dateTerm)
    GetDailyDataTable(term, table)
    
    Dim idx As Integer = 0
    For Each t As DateTerm In term.DailyTerms
      table.Rows(idx)(UserRecordColumnsInfo.DATE_COL_NAME) = f(t)
      idx += 1
    Next
    
    Return table
  End Function
  
  ''' <summary>
  ''' 指定した期間の１日単位のデータを取得する。
  ''' </summary>
  Private Sub GetDailyDataTable(dateTerm As DateTerm, newTable As DataTable)
    If newTable Is Nothing Then Throw New ArgumentNullException("newTable is Null")    
    
    dateTerm.MonthlyTerms.ForEach(
      Sub(term)
        Dim monthlyTable As DataTable = GetRecord(term.BeginDate.Month)
        
        For Each d As DateTerm In term.DailyTerms
          Dim row As Integer = d.BeginDate.Day - 1
          Dim newRow As DataRow = newTable.NewRow
          CopyRow(monthlyTable, row, newRow)
          newTable.Rows.Add(newRow)
        Next
      End Sub)
  End Sub
  
  ''' <summary>
  ''' １列目に日付をつけて指定した期間の１週間単位のデータを取得する。
  ''' </summary>
  Public Function GetWeeklyDataTableLabelingDate(dateTerm As DateTerm, exceptsWorkCountOfZeroWorkTimeIs As Boolean) As DataTable
    Dim table As DataTable = Me.columnsInfo.CreateDataTable(UserRecordColumnsInfo.DATE_COL_NAME)
    Dim term As DateTerm = ModifyDateTerm(dateTerm)
    GetWeeklyDataTable(term, table, exceptsWorkCountOfZeroWorkTimeIs)
    
    Dim weekCountInMonth = DateUtils.GetWeekCountInMonth(term.BeginDate, DayOfWeek.Saturday)
    Dim f As Func(Of DateTime, DateTime, String) =
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
    
    Dim idx As Integer = 0
    For Each t As DateTerm In term.WeeklyTerms(DayOfWeek.Saturday, f)
      table.Rows(idx)(UserRecordColumnsInfo.DATE_COL_NAME) = t.Label
      idx += 1
    Next
    
    Return table
  End Function
  
  ''' <summary>
  ''' 指定した期間の１週間単位のデータを取得する。
  ''' </summary>
  Private Sub GetWeeklyDataTable(dateTerm As DateTerm, newTable As DataTable, exceptsWorkCountOfZeroWorkTimeIs As Boolean)
    If newTable Is Nothing Then Throw New ArgumentNullException("newTable is Null") 
    
    dateTerm.WeeklyTerms.ForEach(
      Sub(w)
        Dim tmpTable As DataTable = Me.columnsInfo.CreateDataTable(String.Empty)
        GetDailyDataTable(w, tmpTable)
      
        Dim newRow As DataRow = newTable.NewRow
        If exceptsWorkCountOfZeroWorkTimeIs Then
          tmpTable.SumExceptWorkCountOfZeroWorkTimeIs(newRow, columnsInfo)          
        Else
          tmpTable.Sum(newRow, Me.columnsInfo)  
        End If
        
        newTable.Rows.Add(newRow)
      ENd Sub)
  End Sub

  ''' <summary>
  ''' １列目に日付をつけて指定した期間の１ヶ月単位のデータを取得する。
  ''' </summary>
  Public Function GetMonthlyDataTableLabelingDate(dateTerm As DateTerm, exceptsWorkCountOfZeroWorkTimeIs As Boolean) As DataTable
    Dim table As DataTable = Me.columnsInfo.CreateDataTable(UserRecordColumnsInfo.DATE_COL_NAME)
    Dim term As DateTerm = ModifyDateTerm(dateTerm)
    GetMonthlyDataTable(term, table, exceptsWorkCountOfZeroWorkTimeIs)
    
    Dim idx As Integer = 0
    For Each t As DateTerm In term.MonthlyTerms(Function(b, e) b.Month & "月")
      table.Rows(idx)(UserRecordColumnsInfo.DATE_COL_NAME) = t.Label
      idx += 1
    Next
    
    Return table
  End Function
  
  ''' <summary>
  ''' 指定した期間の１ヶ月単位のデータを取得する。
  ''' </summary>
  Public Sub GetMonthlyDataTable(dateTerm As DateTerm, newTable As DataTable, exceptsWorkCountOfZeroWorkTimeIs As Boolean)
    If newTable Is Nothing Then Throw New ArgumentNullException("newTable is Null") 
    
     dateTerm.MonthlyTerms.ForEach(
      Sub(m)
        Dim tmpTable As DataTable = Me.columnsInfo.CreateDataTable(String.Empty)
        GetDailyDataTable(m, tmpTable)
      
        Dim newRow As DataRow = newTable.NewRow
        If exceptsWorkCountOfZeroWorkTimeIs Then
          tmpTable.SumExceptWorkCountOfZeroWorkTimeIs(newRow, Me.columnsInfo)
        Else
          tmpTable.Sum(newRow, Me.columnsInfo)  
        End If
        
        newTable.Rows.Add(newRow)
      ENd Sub)   
  End Sub

  ''' <summary>
  ''' 指定した期間の行の集計を返す。
  ''' </summary>
  Public Sub GetTotalDataRow(dateTerm As DateTerm, resultRow As DataRow, exceptsWorkCountOfZeroWorkTimeIs As Boolean)
    Dim table As DataTable = Me.columnsInfo.CreateDataTable(String.Empty)
    Dim term As DateTerm = ModifyDateTerm(dateTerm)
    GetDailyDataTable(term, table)
    
    If exceptsWorkCountOfZeroWorkTimeIs Then
      table.SumExceptWorkCountOfZeroWorkTimeIs(resultRow, Me.columnsInfo)
    Else
      table.Sum(resultRow, Me.columnsInfo)
    End If
  End Sub
  
  ''' <summary>
  ''' 指定したテーブルの指定した行のデータを、別の行コレクションにセットする。
  ''' </summary>
  Private Sub CopyRow(table As DataTable, row As Integer, toRow As DataRow)
    Dim dataRow As DataRow = table.Rows(row)
    For Each column As DataColumn In table.Columns
      toRow(column.ColumnName) = dataRow(column.ColumnName)
    Next
  End Sub
  
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
  
  Private Function ToStringFromFirstColumnItemType(type As UserRecordFirstColumnItemType) As String
    If type = UserRecordFirstColumnItemType.UserName Then
      Return UserRecordColumnsInfo.NAME_COL_NAME
    ElseIf type = UserRecordFirstColumnItemType.DataDate
      Return UserRecordColumnsInfo.DATE_COL_NAME
    Else
      Return String.Empty
    End If
  End Function
  
End Class

Public Enum UserRecordFirstColumnItemType
  UserName
  DataDate
  None
End Enum