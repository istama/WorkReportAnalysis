'
' 日付: 2016/10/25
'
Imports System.Collections.Concurrent
Imports System.Data
Imports System.Linq
Imports Common.Account
Imports Common.Util
Imports Common.Extensions
Imports Common.IO

''' <summary>
''' ユーザレコードを格納しておくクラス。
''' </summary>
Public Class UserRecordBuffer
  Private ReadOnly properties As ExcelProperties
  
  ''' 全ユーザのレコードを格納するテーブル
  Private ReadOnly userRecordDictionary As New ConcurrentDictionary(Of String, UserRecord)
  ''' 全ユーザのレコードの値を日付ごとに集計したテーブルを月ごとに格納したテーブル
  Private ReadOnly totalRecord As UserRecord
  ''' 集計テーブルに加算したユーザのリスト
  Private ReadOnly addedUserListToTotalRecord As New List(Of String)
  
  Public Sub New(properties As ExcelProperties)
    If properties Is Nothing Then Throw New ArgumentNullException("properties is null")
    
    Me.properties = properties
    Me.totalRecord = New UserRecord(New UserInfo("total", "999", "xxx"), Me.properties)
  End Sub
  
  ''' <summary>
  ''' 指定したユーザのレコードを作成する。
  ''' このバッファには登録されない。
  ''' </summary>
  Public Function CreateUserRecord(userInfo As UserInfo) As UserRecord
    If userInfo   Is Nothing Then Throw New ArgumentNullException("userInfo is null")
    
    Return New UserRecord(userInfo, Me.properties)
  End Function

  ''' <summary>
  ''' 指定したユーザのレコードを取得する。
  ''' 存在しない場合は新たに作成し、このバッファに登録する。
  ''' </summary>
  Public Function GetUserRecord(userInfo As UserInfo) As UserRecord
    If userInfo Is Nothing Then Throw New ArgumentNullException("userInfo is null")
    
    Dim record As UserRecord = Nothing
    If Not Me.userRecordDictionary.TryGetValue(userInfo.GetSimpleId, record) Then
      record = CreateUserRecord(userInfo)
      Me.userRecordDictionary.TryAdd(userInfo.GetSimpleId, record)
    End If
    
    Return record
  End Function
  
  ''' <summary>
  ''' すべてのユーザのレコードを取得する。
  ''' </summary>
  ''' <returns></returns>
  Public Function GetUserRecordAll() As List(Of UserRecord)
    Dim list As New List(Of UserRecord)
    
    Dim idArray As String() = Me.userRecordDictionary.Keys.ToArray
    Array.Sort(idArray)
    
    Array.ForEach(
      idArray,
      Sub(id)
        Dim record As UserRecord = Nothing
        If Me.userRecordDictionary.TryGetValue(id, record) Then
          list.Add(record)
        End If
      End Sub)
    
    Return list
  End Function
  
  ''' <summary>
  ''' 集計レコードを取得する。
  ''' </summary>
  Public Function GetTotalRecord() As UserRecord
    Return Me.totalRecord
  End Function
  
  ''' <summary>
  ''' 集計レコードに加算する。
  ''' 無事、加算できた場合はTrueを返す。
  ''' すでに加算されたユーザなら何もせずにFalseを返す。
  ''' </summary>
  Public Function PlusToTotalRecord(userInfo As UserInfo) As Boolean
    SyncLock Me.addedUserListToTotalRecord
      If Not Me.addedUserListToTotalRecord.Contains(userInfo.GetSimpleId) Then
        UpdateTotalRecord(userInfo, Sub(totalRow, userRow) totalRow.PlusByDouble(userRow))
        Me.addedUserListToTotalRecord.Add(userInfo.GetSimpleId)
        Return True
      Else
        Return False
      End If
    End SyncLock
  End Function
  
  ''' <summary>
  ''' 集計レコードから減算する。
  ''' 無事、減算できた場合はTureを返す。
  ''' まだ加算されていないユーザの場合は何もせずFalseを返す。
  ''' </summary>
  Public Function MinusToTotalRecord(userInfo As UserInfo) As Boolean
    SyncLock Me.addedUserListToTotalRecord
      If Me.addedUserListToTotalRecord.Contains(userInfo.GetSimpleId) Then
        UpdateTotalRecord(userInfo, Sub(totalRow, userRow) totalRow.MinusByDouble(userRow))
        Me.addedUserListToTotalRecord.Remove(userInfo.GetSimpleId)
        Return True
      Else
        Return False      
      End If
    End SyncLock
  End Function
  
  ''' <summary>
  ''' 集計レコードを更新する。
  ''' </summary>
  Private Sub UpdateTotalRecord(userInfo As UserInfo, f As Action(Of DataRow, DataRow))
    If userInfo Is Nothing Then Throw New ArgumentNullException("userInfo is null")
    
    Dim record As UserRecord = Nothing
    If Me.userRecordDictionary.TryGetValue(userInfo.GetSimpleId, record) Then
      record.GetRecordDateTerm.MonthlyTerms.ForEach(
        Sub(term)
          Dim d As DateTime = term.BeginDate
          Dim userTable  As DataTable = record.GetDailyDataTable(d.Year, d.Month)
          Dim totalTable As DataTable = Me.totalRecord.GetRecord(d.Month)
          
          For idx As Integer = term.BeginDate.Day - 1 To term.EndDate.Day - 1
            Dim totalRow = totalTable.Rows(idx)
            f(totalRow, userTable.Rows(idx))
            'Log.out("user : " & userInfo.GetSimpleId & " idx: " & idx.ToString & " total Row: " & totalRow.Item(1).ToString) 
          Next
        End Sub)
    Else
      Throw New ArithmeticException("指定したユーザのレコードは存在しません。 / userInfo: " & userInfo.GetName)
    End If    
  End Sub
End Class
