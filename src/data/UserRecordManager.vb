'
' 日付: 2016/10/18
'
Imports System.Data
Imports System.Linq
Imports System.Collections.Concurrent
Imports Common.Account
Imports Common.Util
Imports Common.IO

''' <summary>
''' ユーザレコードを一元管理し、アプリケーション上に表示する表形式に加工して返す。
''' </summary>
Public Class UserRecordManager
  Private ReadOnly properties As ExcelProperties
  
  Private ReadOnly userRecordBuffer As UserRecordBuffer
  Private ReadOnly userRecordLoader As UserRecordLoader
  
  Private ReadOnly unionUser As New UserInfo("union", "xxx", "xxx")
  
  Public Sub New(properties As ExcelProperties)
    If properties Is Nothing Then Throw New ArgumentNullException("properties is null")
    
    Me.properties = properties
    Me.userRecordBuffer = New UserRecordBuffer(properties)
    Me.userRecordLoader = New UserRecordLoader(properties, Me.userRecordBuffer)
  End Sub
  
  Public Function Loader() As UserRecordLoader
    Return Me.userRecordLoader
  End Function
  
  ''' <summary>
  ''' 指定したユーザの指定した月のレコードを取得する。
  ''' </summary>
  Public Function GetDailyRecordLabelingDate(userInfo As UserInfo, year As Integer, month As Integer) As DataTable
    If userInfo Is Nothing Then Throw New ArgumentNullException("userInfo is null")
    Dim record As UserRecord = Me.userRecordBuffer.GetUserRecord(userInfo)
    
    Return record.GetDailyDataTable(year, month)
  End Function
  
  ''' <summary>
  ''' 指定したユーザの指定した期間のレコードを１週間単位で取得する。
  ''' </summary>
  Public Function GetWeeklyRecordLabelingDate(userInfo As UserInfo, dateTerm As DateTerm) As DataTable
    If UserInfo Is Nothing Then Throw New ArgumentNullException("userInfo is null")
    Dim record As UserRecord = Me.userRecordBuffer.GetUserRecord(userInfo)
    
    Return record.GetWeeklyDataTableLabelingDate(dateTerm)
  End Function
  
  ''' <summary>
  ''' 指定したユーザの指定した期間のレコードを１ヶ月単位で取得する。
  ''' </summary>
  Public Function GetMonthlyRecordLabelingDate(userInfo As UserInfo, dateTerm As DateTerm) As DataTable
    If UserInfo Is Nothing Then Throw New ArgumentNullException("userInfo is null")
    Dim record As UserRecord = Me.userRecordBuffer.GetUserRecord(userInfo)
    
    Return record.GetMonthlyDataTableLabelingDate(dateTerm) 
  End Function
  
  ''' <summary>
  ''' 各ユーザごとの指定した期間の集計レコードを集めたテーブルを取得する。
  ''' </summary>
  Public Function GetTallyRecordOfEachUser(dateTerm As DateTerm) As DataTable
    Dim unionRecord As UserRecord = Me.userRecordBuffer.CreateUserRecord(Me.unionUser)
    Dim newTable As DataTable = unionRecord.CreateDataTable(UserRecord.NAME_COL_NAME)
    
    Me.userRecordBuffer.GetUserRecordAll.ForEach(
      Sub(record)
        Dim newRow As DataRow = newTable.NewRow 
        
        newRow(UserRecord.NAME_COL_NAME) = record.GetIdNumber & " " & record.GetName
        record.GetTotalDataRow(dateTerm, newRow)
        newTable.Rows.Add(newRow)
      End Sub)
    
    Return newTable
  End Function
  
  Public Function GetTotalOfAllUserDailyRecord(dateTerm As DateTerm) As DataTable
    Dim totalRecord As UserRecord = Me.userRecordBuffer.GetTotalRecord
    
    Return totalRecord.GetDailyDataTableLabelingDate(dateTerm, Function(t) t.BeginDate.Day & "日")
  End Function
  
  Public Function GetTotalOfAllUserWeeklyRecord(dateTerm As DateTerm) As DataTable
    Dim totalRecord As UserRecord = Me.userRecordBuffer.GetTotalRecord
    
    Return totalRecord.GetWeeklyDataTableLabelingDate(dateTerm)
  End Function
  
  Public Function GetTotalOfAllUserMonthlyRecord(dateTerm As DateTerm) As DataTable
    Dim totalRecord As UserRecord = Me.userRecordBuffer.GetTotalRecord
    
    Return totalRecord.GetMonthlyDataTableLabelingDate(dateTerm)    
  End Function
End Class
