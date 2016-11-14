'
' 日付: 2016/10/18
'
Imports System.Threading.Tasks
Imports System.Collections.Concurrent

Imports Common.Account
Imports Common.COM
Imports Common.IO

Public Class UserRecordLoader
  Private ReadOnly excel As Excel3
  
  Private ReadOnly userRecordBuffer As UserRecordBuffer 
  Private ReadOnly userRecordReader As UserRecordReader
  
  Private ReadOnly tasks As New ConcurrentQueue(Of Task)
  
  Public Sub New(properties As ExcelProperties, userRecordBuffer As UserRecordBuffer)
    If properties Is Nothing Then Throw New ArgumentNullException("properties is null")
    If userRecordBuffer Is Nothing Then Throw New ArgumentNullException("userRecordBuffer is null ")
    
    Me.excel = New Excel3()
    Me.userRecordBuffer = userRecordBuffer
    Me.userRecordReader = New UserRecordReader(properties, Me.excel)
  End Sub
  
  Public Sub Init()
    'Me.excel.init
  End Sub
  
  Public Sub Quit()
    'Me.excel.Quit
  End Sub
  
  Public Sub Wait
    While True
      Threading.Thread.Sleep(50)
      If Task.WaitAll(Me.tasks.ToArray, 0) Then
        Exit While
      End If
    End While
  End Sub
  
  Public Sub Load(loadedUserInfo As UserInfo, f As Action(Of UserRecord))
    'tasks.Enqueue(Task.Factory.StartNew(Sub() LoadTask(loadedUserInfo, f)))
    LoadTask(loadedUserInfo, f)
  End Sub
  
  Private Sub LoadTask(loadedUserInfo As UserInfo, f As Action(Of UserRecord))
    Try
      Dim record As UserRecord = Me.userRecordBuffer.GetUserRecord(loadedUserInfo)
      Me.userRecordBuffer.MinusToTotalRecord(loadedUserInfo)
      Me.userRecordReader.Read(record)
      Me.userRecordBuffer.PlusToTotalRecord(loadedUserInfo)
      f(record)
    Catch ex As Exception
      Log.out(ex.Message)
    End Try
  End Sub
  
  Public Sub Cancel()
    Me.userRecordReader.Cancel
  End Sub
End Class
