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
  Private ReadOnly properties As ExcelProperties
  
  Private ReadOnly recordColumnsInfo As UserRecordColumnsInfo
  
  Private ReadOnly userRecordBuffer As UserRecordBuffer
  Private ReadOnly userRecordLoader As UserRecordLoader
  
  Private ReadOnly unionUser As New UserInfo("union", "xxx", "xxx")
  
  Public Sub New(properties As ExcelProperties)
    If properties Is Nothing Then Throw New ArgumentNullException("properties is null")
    
    Me.properties = properties
    Me.recordColumnsInfo = New UserRecordColumnsInfo(properties)
    Me.userRecordBuffer = New UserRecordBuffer(properties)
    Me.userRecordLoader = New UserRecordLoader(properties, Me.userRecordBuffer)
  End Sub
  
  Public Function Loader() As UserRecordLoader
    Return Me.userRecordLoader
  End Function
  
  Public Function GetUserRecordColumnsInfo() As UserRecordColumnsInfo
    Return Me.recordColumnsInfo
  End Function
  
  ''' <summary>
  ''' 指定したユーザの指定した月のレコードを取得する。
  ''' </summary>
  Public Function GetDailyRecordLabelingDate(userInfo As UserInfo, year As Integer, month As Integer) As DataTable
    If userInfo Is Nothing Then Throw New ArgumentNullException("userInfo is null")
    Dim record As UserRecord = Me.userRecordBuffer.GetUserRecord(userInfo)
    
    Dim table As DataTable = record.GetDailyDataTable(year, month) 
    AddTotalRowToTable(table, UserRecordColumnsInfo.DATE_COL_NAME)
    
    Return AddColumnOfProductivityToRecord(table)
    'Return record.GetDailyDataTable(year, month)
  End Function
  
  ''' <summary>
  ''' 指定したユーザの指定した期間のレコードを１週間単位で取得する。
  ''' </summary>
  Public Function GetWeeklyRecordLabelingDate(userInfo As UserInfo, dateTerm As DateTerm) As DataTable
    If UserInfo Is Nothing Then Throw New ArgumentNullException("userInfo is null")
    Dim record As UserRecord = Me.userRecordBuffer.GetUserRecord(userInfo)
    
    Dim table As DataTable = record.GetWeeklyDataTableLabelingDate(dateTerm)
    AddTotalRowToTable(table, UserRecordColumnsInfo.DATE_COL_NAME)
    
    Return AddColumnOfProductivityToRecord(table)
  End Function
  
  ''' <summary>
  ''' 指定したユーザの指定した期間のレコードを１ヶ月単位で取得する。
  ''' </summary>
  Public Function GetMonthlyRecordLabelingDate(userInfo As UserInfo, dateTerm As DateTerm) As DataTable
    If UserInfo Is Nothing Then Throw New ArgumentNullException("userInfo is null")
    Dim record As UserRecord = Me.userRecordBuffer.GetUserRecord(userInfo)
    
    Dim table As DataTable = record.GetMonthlyDataTableLabelingDate(dateTerm)
    AddTotalRowToTable(table, UserRecordColumnsInfo.DATE_COL_NAME)
    
    Return AddColumnOfProductivityToRecord(table)
  End Function
  
  ''' <summary>
  ''' 指定したユーザの集計レコードを取得する。
  ''' </summary>
  Public Function GetSumRecord(userInfo As UserInfo) As DataTable
    If userInfo Is Nothing Then Throw New ArgumentNullException("userInfo is null")
    Dim record As UserRecord = Me.userRecordBuffer.GetUserRecord(userInfo)
    
    ' 1週間ごとの集計テーブルを取得
    Dim weeklySumTable As DataTable  = record.GetWeeklyDataTableLabelingDate(record.GetRecordDateTerm)
    Dim monthlySumTable As DataTable = record.GetMonthlyDataTableLabelingDate(record.GetRecordDateTerm)
    AddTotalRowToTable(monthlySumTable, UserRecordColumnsInfo.DATE_COL_NAME)
    
    ' weeklySumTableの後ろの行にmonthlySumTableの行を追加する
    For Each row As DataRow In monthlySumTable.Rows
      Dim newRow As DataRow = weeklySumTable.NewRow
      For Each col As DataColumn In weeklySumTable.Columns
        newRow(col.ColumnName) = row(col.ColumnName)
      Next
      
      weeklySumTable.Rows.Add(newRow)
    Next
    
    Return AddColumnOfProductivityToRecord(weeklySumTable)
  End Function
  
  ''' <summary>
  ''' 各ユーザごとの指定した期間の集計レコードを集めたテーブルを取得する。
  ''' </summary>
  Public Function GetTallyRecordOfEachUser(dateTerm As DateTerm) As DataTable
    Dim unionRecord As UserRecord = Me.userRecordBuffer.CreateUserRecord(Me.unionUser)
    Dim newTable As DataTable = unionRecord.CreateDataTable(UserRecordColumnsInfo.NAME_COL_NAME)
    
    Me.userRecordBuffer.GetUserRecordAll.ForEach(
      Sub(record)
        Dim newRow As DataRow = newTable.NewRow 
        
        newRow(UserRecordColumnsInfo.NAME_COL_NAME) = record.GetIdNumber & " " & record.GetName
        record.GetTotalDataRow(dateTerm, newRow)
        newTable.Rows.Add(newRow)
      End Sub)
    
    AddTotalRowToTable(newTable, UserRecordColumnsInfo.NAME_COL_NAME)
    
    Return AddColumnOfProductivityToRecord(newTable)
  End Function
  
  Public Function GetTotalOfAllUserDailyRecord(dateTerm As DateTerm) As DataTable
    Dim totalRecord As UserRecord = Me.userRecordBuffer.GetTotalRecord
    
    Dim table As DataTable = totalRecord.GetDailyDataTableLabelingDate(dateTerm, Function(t) t.BeginDate.Day & "日")
    AddTotalRowToTable(table, UserRecordColumnsInfo.DATE_COL_NAME)
    
    Return AddColumnOfProductivityToRecord(table)
  End Function
  
  Public Function GetTotalOfAllUserWeeklyRecord(dateTerm As DateTerm) As DataTable
    Dim totalRecord As UserRecord = Me.userRecordBuffer.GetTotalRecord
    
    Dim table As DataTable = totalRecord.GetWeeklyDataTableLabelingDate(dateTerm)
    AddTotalRowToTable(table, UserRecordColumnsInfo.DATE_COL_NAME) 
    
    Return AddColumnOfProductivityToRecord(table)
  End Function
  
  Public Function GetTotalOfAllUserMonthlyRecord(dateTerm As DateTerm) As DataTable
    Dim totalRecord As UserRecord = Me.userRecordBuffer.GetTotalRecord
    
    Dim table As DataTable = totalRecord.GetMonthlyDataTableLabelingDate(dateTerm)
    AddTotalRowToTable(table, UserRecordColumnsInfo.DATE_COL_NAME)
    
    Return AddColumnOfProductivityToRecord(table)
  End Function
  
  ''' <summary>
  ''' 合計行を追加する。
  ''' </summary>
  Private Sub AddTotalRowToTable(table As DataTable, firstColumnName As String)
    Dim totalRow As DataRow = table.NewRow
    totalRow(firstColumnName) = "合計"
    For Each row As DataRow In table.Rows
      totalRow.PlusByDouble(row)
    Next
    table.Rows.Add(totalRow)
  End Sub
  
  ''' <summary>
  ''' ユーザデータに生産性の列を追加して返す。
  ''' </summary>
  Private Function AddColumnOfProductivityToRecord(table As DataTable) As DataTable
    Dim newTable As New DataTable
    
    ' 新しいテーブルに「名前」もしくは「日にち」の列を追加
    newTable.Columns.Add(table.Columns(0).ColumnName)
    
    ' 新しいテーブルに作業項目の列を追加
    Dim workItemCnt As Integer = 1
    For Each item As WorkItemColumnsInfo In Me.recordColumnsInfo.WorkItemList
      Dim isThereCntColName As Boolean = False
      Dim isThereTimeColName As Boolean = False
      If Not String.IsNullOrWhiteSpace(item.WorkCountColName) Then
        Log.out(item.WorkCountColName)
        newTable.Columns.Add(item.WorkCountColName)
        isThereCntColName = True
      End If
      If Not String.IsNullOrWhiteSpace(item.WorkTimeColName) Then
        Log.out(item.WorkTimeColName)
        newTable.Columns.Add(item.WorkTimeColName)
        isThereTimeColName = True
      End If
      If isThereCntColName AndAlso isThereTimeColName Then
        Log.out(item.WorkProductivityColName)
        newTable.Columns.Add(item.WorkProductivityColName)
      End If
    Next
    
    ' 新しいテーブルに備考の列を追加
    newTable.Columns.Add(Me.recordColumnsInfo.noteColName)
    
    ' データをコピー
    For Each row As DataRow In table.Rows
      Dim newRow As DataRow = newTable.NewRow
      For Each col As DataColumn In table.Columns
        newRow(col.ColumnName) = row(col.ColumnName)
      Next
      
      ' 生産性を計算
      For Each item As WorkItemColumnsInfo In Me.recordColumnsInfo.WorkItemList
        Dim cntColName As String  = item.WorkCountColName
        Dim timeColName As String = item.WorkTimeColName
        If cntColName <> String.Empty AndAlso timeColName <> String.Empty Then
          Dim cntValue As Object = row(cntColName)
          Dim timeValue As Object = row(timeColName)
          If Not System.Convert.IsDBNull(cntValue) AndAlso Not System.Convert.IsDBNull(timeValue) Then
            Dim cnt As Double
            Dim time As Double
            If Double.TryParse(DIrectCast(cntValue, String), cnt) AndAlso Double.TryParse(DirectCast(timeValue, String), time) Then
              If time <> 0 Then
                Dim productivity As Double = cnt / time
                newRow(item.WorkProductivityColName) = Math.Round(productivity, 2).ToString
              End If
            End If
          End If
        End If
      Next
      
      newTable.Rows.Add(newRow)
    Next
    
    Return newTable
  End Function
End Class
