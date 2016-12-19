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
    Me.userRecordLoader = New UserRecordLoader(properties)
    Me.userRecordBuffer = Me.userRecordLoader.GetUserRecordBuffer()
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
  Public Function GetDailyRecordLabelingDate(userInfo As UserInfo, year As Integer, month As Integer) As DataTable
    If userInfo Is Nothing Then Throw New ArgumentNullException("userInfo is null")
    
    Dim record As UserRecord = Me.userRecordBuffer.GetUserRecord(userInfo)
    Dim table As DataTable = record.GetDailyDataTableForAMonth(year, month) 
    
    AddTotalRowToTable(table, UserRecordColumnsInfo.DATE_COL_NAME)
    
    Return AddColumnOfProductivityToRecord(table)
  End Function
  
  ''' <summary>
  ''' 指定したユーザの指定した期間のレコードを１週間単位で取得する。
  ''' </summary>
  Public Function GetWeeklyRecordLabelingDate(userInfo As UserInfo, dateTerm As DateTerm) As DataTable
    If UserInfo Is Nothing Then Throw New ArgumentNullException("userInfo is null")
    Dim record As UserRecord = Me.userRecordBuffer.GetUserRecord(userInfo)
    
    Dim table As DataTable = record.GetWeeklyDataTableLabelingDate(dateTerm, False)
    AddTotalRowToTable(table, UserRecordColumnsInfo.DATE_COL_NAME)
    
    Return AddColumnOfProductivityToRecord(table)
  End Function
  
  ''' <summary>
  ''' 指定したユーザの指定した期間のレコードを１ヶ月単位で取得する。
  ''' </summary>
  Public Function GetMonthlyRecordLabelingDate(userInfo As UserInfo, dateTerm As DateTerm) As DataTable
    If UserInfo Is Nothing Then Throw New ArgumentNullException("userInfo is null")
    Dim record As UserRecord = Me.userRecordBuffer.GetUserRecord(userInfo)
    
    Dim table As DataTable = record.GetMonthlyDataTableLabelingDate(dateTerm, False)
    AddTotalRowToTable(table, UserRecordColumnsInfo.DATE_COL_NAME)
    
    Return AddColumnOfProductivityToRecord(table)
  End Function
  
  ''' <summary>
  ''' 指定したユーザの集計レコードを取得する。
  ''' </summary>
  Public Function GetSumRecord(userInfo As UserInfo, exceptsWorkCountOfZeroWorkTimeIs As Boolean) As DataTable
    If userInfo Is Nothing Then Throw New ArgumentNullException("userInfo is null")
    Dim record As UserRecord = Me.userRecordBuffer.GetUserRecord(userInfo)
    
    ' 1週間ごとの集計テーブルを取得
    Dim weeklySumTable As DataTable  = record.GetWeeklyDataTableLabelingDate(record.GetRecordDateTerm, exceptsWorkCountOfZeroWorkTimeIs)
    Dim monthlySumTable As DataTable = record.GetMonthlyDataTableLabelingDate(record.GetRecordDateTerm, exceptsWorkCountOfZeroWorkTimeIs)
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
  Public Function GetTallyRecordOfEachUser(dateTerm As DateTerm, exceptsWorkCountOfZeroWorkTimeIs As Boolean) As DataTable
    Dim unionRecord As UserRecord = Me.userRecordBuffer.CreateUserRecord(Me.unionUser)
    Dim newTable As DataTable = Me.recordColumnsInfo.CreateDataTable(UserRecordColumnsInfo.NAME_COL_NAME)
    
    Me.userRecordBuffer.GetUserRecordAll.ForEach(
      Sub(record)
        Dim newRow As DataRow = newTable.NewRow 
        
        newRow(UserRecordColumnsInfo.NAME_COL_NAME) = record.GetIdNumber & " " & record.GetName
        record.GetTotalDataRow(dateTerm, newRow, exceptsWorkCountOfZeroWorkTimeIs)
        newTable.Rows.Add(newRow)
      End Sub)
    
    AddTotalRowToTable(newTable, UserRecordColumnsInfo.NAME_COL_NAME)
    
    Return AddColumnOfProductivityToRecord(newTable)
  End Function
  
  Public Function GetTotalOfAllUserDailyRecord(dateTerm As DateTerm, exceptsWorkCountOfZeroWorkTimeIs As Boolean) As DataTable
    Dim totalRecord As UserRecord
    If exceptsWorkCountOfZeroWorkTimeIs Then
      totalRecord = Me.userRecordBuffer.GetTotalRecordExceptedWorkCountOfZeroWorkTimeIs
    Else
      totalRecord = Me.userRecordBuffer.GetTotalRecord  
    End If
    
    Dim table As DataTable = totalRecord.GetDailyDataTableLabelingDate(dateTerm, Function(t) t.BeginDate.Day & "日")
    AddTotalRowToTable(table, UserRecordColumnsInfo.DATE_COL_NAME)
    
    Return AddColumnOfProductivityToRecord(table)
  End Function
  
  Public Function GetTotalOfAllUserWeeklyRecord(dateTerm As DateTerm, exceptsWorkCountOfZeroWorkTimeIs As Boolean) As DataTable
    Dim totalRecord As UserRecord
    If exceptsWorkCountOfZeroWorkTimeIs Then
      totalRecord = Me.userRecordBuffer.GetTotalRecordExceptedWorkCountOfZeroWorkTimeIs
    Else
      totalRecord = Me.userRecordBuffer.GetTotalRecord  
    End If
    
    Dim table As DataTable = totalRecord.GetWeeklyDataTableLabelingDate(dateTerm, False)
    AddTotalRowToTable(table, UserRecordColumnsInfo.DATE_COL_NAME) 
    
    Return AddColumnOfProductivityToRecord(table)
  End Function
  
  Public Function GetTotalOfAllUserMonthlyRecord(dateTerm As DateTerm, exceptsWorkCountOfZeroWorkTimeIs As Boolean) As DataTable
    Dim totalRecord As UserRecord
    If exceptsWorkCountOfZeroWorkTimeIs Then
      totalRecord = Me.userRecordBuffer.GetTotalRecordExceptedWorkCountOfZeroWorkTimeIs
    Else
      totalRecord = Me.userRecordBuffer.GetTotalRecord  
    End If
    
    Dim table As DataTable = totalRecord.GetMonthlyDataTableLabelingDate(dateTerm, False)
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
    Dim newTable As DataTable = Me.recordColumnsInfo.CreateDataTable(table.Columns(0).ColumnName)

    ' データをコピー
    For Each row As DataRow In table.Rows
      Dim newRow As DataRow = newTable.NewRow
      For Each col As DataColumn In table.Columns
        newRow(col.ColumnName) = row(col.ColumnName)
      Next
      
      ' 生産性を計算
      For Each item As WorkItemColumnsInfo In Me.recordColumnsInfo.WorkItemList
        Dim cntColName As String  = item.WorkCountColInfo.Name
        Dim timeColName As String = item.WorkTimeColInfo.Name
        If cntColName <> String.Empty AndAlso timeColName <> String.Empty Then
          Dim cntValue As Object = row(cntColName)
          Dim timeValue As Object = row(timeColName)
          If Not System.Convert.IsDBNull(cntValue) AndAlso Not System.Convert.IsDBNull(timeValue) Then
            Dim cnt As Integer = DirectCast(cntValue, Integer)
            Dim time As Double = DirectCast(timeValue, Double)
            'If Double.TryParse(DIrectCast(cntValue, String), cnt) AndAlso Double.TryParse(DirectCast(timeValue, String), time) Then
              If time <> 0 Then
                Dim productivity As Double = cnt / time
                newRow(item.WorkProductivityColInfo.Name) = Math.Round(productivity, 2)
              End If
            'End If
          End If
        End If
      Next
      
      newTable.Rows.Add(newRow)
    Next
    
    Return newTable
  End Function
End Class
