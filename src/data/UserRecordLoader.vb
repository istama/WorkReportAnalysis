'
' 日付: 2016/10/18
'
Imports System.Threading.Tasks
Imports System.Collections.Concurrent

Imports Common.Account
Imports Common.COM
Imports Common.IO

Public Class UserRecordLoader
  Private ReadOnly excel As Excel4
  
  Private ReadOnly userRecordBuffer As UserRecordBuffer 
  Private ReadOnly userRecordReader As UserRecordReader
  
  Private ReadOnly tasks As New ConcurrentQueue(Of Task)
  
  Public Sub New(properties As ExcelProperties)
    If properties Is Nothing Then Throw New ArgumentNullException("properties is null")
    
    Me.excel = New Excel4()
    Me.userRecordBuffer = New UserRecordBuffer(properties)
    Me.userRecordReader = New UserRecordReader(properties, Me.excel)
  End Sub
  
  Public Sub Init()
    #If Debug = False Then
      Me.excel.init
    #End If
  End Sub
  
  Public Sub Quit()
    #If Debug = False Then
      Me.excel.Quit
    #End If
  End Sub
  
  Public Sub Wait
    While True
      Threading.Thread.Sleep(50)
      If Task.WaitAll(Me.tasks.ToArray, 0) Then
        Exit While
      End If
    End While
  End Sub
  
  Public Sub Load(loadedUserInfo As UserInfo)
    Load(loadedUserInfo, Nothing)
  End Sub
  
  Public Sub Load(loadedUserInfo As UserInfo, callback As Action(Of UserRecord))
    'tasks.Enqueue(Task.Factory.StartNew(Sub() LoadTask(loadedUserInfo, f)))
    LoadTask(loadedUserInfo, callback)
  End Sub
  
  Private Sub LoadTask(loadedUserInfo As UserInfo, callback As Action(Of UserRecord))
    Try
      Dim record As UserRecord = Me.userRecordBuffer.GetUserRecord(loadedUserInfo)
      Me.userRecordBuffer.MinusToTotalRecord(loadedUserInfo)
      Me.userRecordReader.Read(record)
      Me.userRecordBuffer.PlusToTotalRecord(loadedUserInfo)
      If callback IsNot Nothing Then
        callback(record)
      End If
    Catch ex As Exception
      Log.out(ex.Message)
    End Try
  End Sub
  
  Public Sub Cancel()
    Me.userRecordReader.Cancel
  End Sub
End Class
