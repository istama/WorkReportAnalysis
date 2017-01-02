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
  
  Private ReadOnly dataDateTerm As DateTerm
  
  Public Sub New(properties As ExcelProperties)
    If properties Is Nothing Then Throw New ArgumentNullException("properties is null")
    
    Me.properties        = properties
    Me.recordColumnsInfo = UserRecordColumnsInfo.Create(properties)
    Me.dataDateTerm      = New DateTerm(properties.BeginDate, properties.EndDate) 
    
    ' 全ユーザの集計用レコードを作成
    Dim dummyUser As New UserInfo("dummy", "999", "xxx") ' このパラメータに意味はない
    Me.totalRecord                 = New UserRecord(dummyUser, Me.recordColumnsInfo, Me.dataDateTerm)
    Me.totalRecordExceptedUnfilled = New UserRecord(dummyUser, Me.recordColumnsInfo, Me.dataDateTerm)
  End Sub
  
  ''' <summary>
  ''' 指定したユーザのレコードを作成する。
  ''' このバッファには登録されない。
  ''' </summary>
  Public Function CreateUserRecord(userInfo As UserInfo) As UserRecord
    If userInfo   Is Nothing Then Throw New ArgumentNullException("userInfo is null")
    
    Return New UserRecord(userInfo, Me.recordColumnsInfo, Me.dataDateTerm)
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
  Public Function GetTotalRecordExceptedUnfilled() As UserRecord
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
        ' 集計レコードにユーザレコードの値を加算する
        UpdateRecord(
          userInfo,
          Me.totalRecord,
          Sub(userTable, updatedTable) PlusAllData(updatedTable, userTable))
        
        ' 集計レコードにユーザレコードの値を加算する
        ' ただし、作業時間が０の作業件数は除く
        UpdateRecord(
          userInfo,
          Me.totalRecordExceptedUnfilled,
          Sub(userTable, updatedTable) PlusAllDataExceptedUnfilled(updatedTable, userTable))
        
        Me.addedUserListToTotalRecord.Add(userInfo.GetSimpleId)
        Return True
      Else
        Return False
      End If
    End SyncLock
  End Function
  
  ''' <summary>
  ''' 集計レコードから減算する。
  ''' 無事、減算できた場合はTrueを返す。
  ''' 存在しないユーザなら何もせずにFalseを返す。
  ''' </summary>
  Public Function MinusToTotalRecord(userInfo As UserInfo) As Boolean
    SyncLock Me.addedUserListToTotalRecord
      If Me.addedUserListToTotalRecord.Contains(userInfo.GetSimpleId) Then
        ' 集計レコードからユーザレコードの値を減算する
        UpdateRecord(
          userInfo,
          Me.totalRecord,
          Sub(userTable, updatedTable) MinusAllData(updatedTable, userTable))
        
        ' 集計レコードからユーザレコードの値を減算する
        ' ただし、作業時間が０の作業件数は除く
        UpdateRecord(
          userInfo,
          Me.totalRecordExceptedUnfilled,
          Sub(userTable, updatedTable) MinusAllDataExceptedUnfilled(updatedTable, userTable))
        
        Me.addedUserListToTotalRecord.Remove(userInfo.GetSimpleId)
        Return True
      Else
        Return False
      End If
    End SyncLock
  End Function
  
  ''' <summary>
  ''' 同じ列、同じ行にあるデータ同士を加算する。
  ''' </summary>
  Private Sub PlusAllData(addedTable As DataTable, addingTable As DataTable)
    Dim size As Integer = Math.Min(addedTable.Rows.Count, addingTable.Rows.Count)
    
    For idx As Integer = 0 To size - 1
      addedTable.Rows(idx).Plus(addingTable.Rows(idx))
    Next
  End Sub
  
  ''' <summary>
  ''' 同じ列、同じ行にあるデータ同士を加算する。
  ''' ただし、作業時間が０の作業件数は加算しない。
  ''' </summary>
  Private Sub PlusAllDataExceptedUnfilled(addedTable As DataTable, addingTable As DataTable)
    Dim size As Integer = Math.Min(addedTable.Rows.Count, addingTable.Rows.Count)
    
    ' 各行の値を加算する
    For idx As Integer = 0 To size - 1
      Dim addedRow  As DataRow = addedTable.Rows(idx)
      Dim addingRow As DataRow = addingTable.Rows(idx)
      
      For Each workItems As WorkItemColumnsInfo In Me.recordColumnsInfo.WorkItems
        Dim timeColName As String = workItems.WorkTimeColInfo.name
        
        If Not String.IsNullOrWhiteSpace(timeColName) Then
          Dim time As Double = addingRow.GetOrDefault(Of Double)(timeColName, 0.0)
          ' 作業時間が０以上の場合のみ、作業件数を加算する
          If time > 0.0 Then
            addedRow(timeColName) = addedRow.GetOrDefault(Of Double)(timeColName, 0.0) + time
            
            Dim cntColName As String = workItems.WorkCountColInfo.name
            
            If Not String.IsNullOrWhiteSpace(cntColName) Then
              addedRow(cntColName) = addedRow.GetOrDefault(Of Integer)(cntColName, 0) + addingRow.GetOrDefault(Of Integer)(cntColName, 0)
            End If
          End If
        End If
      Next
    Next
  End Sub
  
  ''' <summary>
  ''' 同じ列、同じ行にあるデータ同士を減算する。
  ''' </summary>
  Private Sub MinusAllData(subtractedTable As DataTable, subtractingTable As DataTable)
    Dim size As Integer = Math.Min(subtractedTable.Rows.Count, subtractingTable.Rows.Count)
    
    For idx As Integer = 0 To size - 1
      subtractedTable.Rows(idx).Minus(subtractingTable.Rows(idx))
    Next
  End Sub
  
  ''' <summary>
  ''' 同じ列、同じ行にあるデータ同士を加算する。
  ''' ただし、作業時間が０の作業件数は減算しない。
  ''' </summary>
  Private Sub MinusAllDataExceptedUnfilled(subtractedTable As DataTable, subtractingTable As DataTable)
    Dim size As Integer = Math.Min(subtractedTable.Rows.Count, subtractingTable.Rows.Count)
    
    ' 各行の値を減算する
    For idx As Integer = 0 To size - 1
      Dim subtractedRow  As DataRow = subtractedTable.Rows(idx)
      Dim subtractingRow As DataRow = subtractingTable.Rows(idx)
      
      For Each workItems As WorkItemColumnsInfo In Me.recordColumnsInfo.WorkItems
        Dim timeColName As String = workItems.WorkTimeColInfo.name
        
        If Not String.IsNullOrWhiteSpace(timeColName) Then
          Dim time As Double = subtractingRow.GetOrDefault(Of Double)(timeColName, 0.0)
          ' 作業時間が０以上の場合のみ、作業件数を加算する
          If time > 0.0 Then
            subtractedRow(timeColName) = subtractedRow.GetOrDefault(Of Double)(timeColName, 0.0) - time
            
            Dim cntColName As String  = workItems.WorkCountColInfo.name
            
            If Not String.IsNullOrWhiteSpace(cntColName) Then
              subtractedRow(cntColName) = subtractedRow.GetOrDefault(Of Integer)(cntColName, 0) - subtractingRow.GetOrDefault(Of Integer)(cntColName, 0)
            End If
          End If
        End If
      Next
    Next
  End Sub
  
  ''' <summary>
  ''' 指定したユーザのレコードと指定した更新されるレコードのテーブルを月ごとに呼び出し、
  ''' コールバック関数に引き渡す。
  ''' </summary>
  Private Sub UpdateRecord(userInfo As UserInfo, updatedRecord As UserRecord, update As Action(Of DataTable, DataTable))
    If userInfo      Is Nothing Then Throw New ArgumentNullException("userInfo is null")
    If updatedRecord Is Nothing Then Throw New ArgumentNullException("updatedRecord is null")
    If update        Is Nothing Then Throw NEw ArgumentNullException("update is null")
    
    Dim record As UserRecord = Nothing
    
    If Me.userRecordDictionary.TryGetValue(userInfo.GetSimpleId, record) Then
      For Each term As DateTerm In record.GetRecordDateTerm.MonthlyTerms
        Dim d            As DateTime  = term.BeginDate
        Dim userTable    As DataTable = record.GetRecord(d.Month)
        Dim updatedTable As DataTable = updatedRecord.GetRecord(d.Month)
        update(userTable, updatedTable)
      Next
    Else
      Throw New ArgumentException("指定したユーザのレコードは存在しません。 / userInfo: " & userInfo.GetName)
    End If
  End Sub
End Class
