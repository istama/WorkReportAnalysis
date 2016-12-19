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
  
  ''' レコードの列名
  Private ReadOnly recordColumnsInfo As UserRecordColumnsInfo
  
  ''' 全ユーザのレコードを格納するテーブル
  Private ReadOnly userRecordDictionary As New ConcurrentDictionary(Of String, UserRecord)
  ''' 全ユーザのレコードの値を日付ごとに集計したテーブルを月ごとに格納したテーブル
  Private ReadOnly totalRecord As UserRecord
  ''' 全ユーザのレコードの値を日付ごとに集計したテーブルを月ごとに格納したテーブル
  ''' ただし作業時間が０の作業件数は集計に含めない
  Private ReadOnly totalRecordExceptedUnfilled As UserRecord
  
  ''' 集計テーブルに加算したユーザのリスト
  Private ReadOnly addedUserListToTotalRecord As New List(Of String)
  
  Public Sub New(properties As ExcelProperties)
    If properties Is Nothing Then Throw New ArgumentNullException("properties is null")
    
    Me.properties = properties
    Me.recordColumnsInfo = New UserRecordColumnsInfo(properties)
    
    Me.totalRecord = New UserRecord(New UserInfo("total", "999", "xxx"), Me.recordColumnsInfo, properties)
    Me.totalRecordExceptedUnfilled = New UserRecord(New UserInfo("total", "999", "xxx"), Me.recordColumnsInfo, properties)
  End Sub
  
  ''' <summary>
  ''' 指定したユーザのレコードを作成する。
  ''' このバッファには登録されない。
  ''' </summary>
  Public Function CreateUserRecord(userInfo As UserInfo) As UserRecord
    If userInfo   Is Nothing Then Throw New ArgumentNullException("userInfo is null")
    
    Return New UserRecord(userInfo, Me.recordColumnsInfo, Me.properties)
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
  ''' 指定したユーザが登録されているか判定する。
  ''' </summary>
  Public Function Stored(userInfo As UserInfo) As Boolean
    If userInfo Is Nothing Then Throw New ArgumentNullException("userInfo is null")
    
    Return Me.userRecordDictionary.ContainsKey(userInfo.GetSimpleId)
  End Function
  
  ''' <summary>
  ''' すべてのユーザのレコードを取得する。
  ''' </summary>
  ''' <returns></returns>
  Public Iterator Function GetUserRecordAll() As IEnumerable(Of UserRecord)
    Dim idArray As String() = Me.userRecordDictionary.Keys.ToArray
    Array.Sort(idArray)
    
    For Each id As String In idArray
      Dim record As UserRecord = Nothing
      If Me.userRecordDictionary.TryGetValue(id, record) Then
        Yield record
      End If
    Next
  End Function
  
  ''' <summary>
  ''' 集計レコードを取得する。
  ''' </summary>
  Public Function GetTotalRecord() As UserRecord
    Return Me.totalRecord
  End Function
  
  ''' <summary>
  ''' 作業時間が０の作業件数が除かれた集計レコードを取得する。
  ''' </summary>
  Public Function GetTotalRecordExceptedWorkCountOfZeroWorkTimeIs() As UserRecord
    Return Me.totalRecordExceptedUnfilled
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
        ' TODO マイナスが行われていない
        UpdateTotalRecord(userInfo, Sub(totalRow, userRow) totalRow.MinusByDouble(userRow))
        Me.addedUserListToTotalRecord.Remove(userInfo.GetSimpleId)
        Return True
      Else
        Return False      
      End If
    End SyncLock
  End Function
  
  ''' <summary>
  ''' TODO 見直す
  ''' 集計レコードを更新する。
  ''' </summary>
  Private Sub UpdateTotalRecord(userInfo As UserInfo, f As Action(Of DataRow, DataRow))
    If userInfo Is Nothing Then Throw New ArgumentNullException("userInfo is null")
    
    Dim record As UserRecord = Nothing
    If Me.userRecordDictionary.TryGetValue(userInfo.GetSimpleId, record) Then
      record.GetRecordDateTerm.MonthlyTerms.ForEach(
        Sub(term)
          Dim d As DateTime = term.BeginDate
          Dim userTable  As DataTable = record.GetDailyDataTableForAMonth(d.Year, d.Month)
          Dim totalTable As DataTable = Me.totalRecord.GetRecord(d.Month)
          Dim totalTableE As DataTable = Me.totalRecordExceptedUnfilled.GetRecord(d.Month)
          
          userTable.Plus(totalTable, Me.recordColumnsInfo)
          userTable.PlusExceptingWorkCountOfZerpWorkTimeIs(totalTableE, Me.recordColumnsInfo)
        End Sub)
    Else
      Throw New ArgumentException("指定したユーザのレコードは存在しません。 / userInfo: " & userInfo.GetName)
    End If    
  End Sub
End Class
