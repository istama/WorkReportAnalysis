﻿'
' 日付: 2016/10/18
'
Imports Common.Account
Imports Common.IO

Public Class Loader
  Private userInfoManager As UserInfoManager
  Private userRecordManager As UserRecordManager
  Private userRecordLoader As UserRecordLoader
  
  Public Sub New(userRecordManager As UserRecordManager, userInfoManager As UserInfoManager, properties As ExcelProperties)
    If userRecordManager Is Nothing Then Throw New ArgumentNullException("userREcordManager is null")
    If userInfoManager Is Nothing Then Throw New ArgumentNullException("userInfoManager is null")
    If properties      Is Nothing Then Throw New ArgumentNullException("properties is null")
    
    Me.userInfoManager   = userInfoManager
    Me.userRecordManager = userRecordManager
    Me.userRecordLoader  = userRecordManager.Loader
  End Sub
  
  ''' <summary>
  ''' 読み込むデータ件数を返す。
  ''' </summary>
  Public Function LoadedCount As Integer
    Return Me.UserInfoManager.UserInfoList.Count
  End Function
  
  Public Sub Quit
    Me.userRecordLoader.Quit
  End Sub
  
  ''' <summary>
  ''' 読み込み処理を開始する。
  ''' </summary>
  Public Sub Load(observer As ThreadObserver)
    If observer Is Nothing Then Throw New ArgumentNullException("observer is null")
    
    Me.userRecordLoader.Init
    
    ' 全ユーザのレコードを読み込む
    Me.userInfoManager.UserInfoList.ForEach(
      Sub(ui)
        Me.userRecordLoader.Load(
          ui,
          Sub(record) observer.ReportProgress(record.GetIdNumber & " " & record.GetName)
        )
      End Sub)
    
    'Log.out("Loader: wait")
    Me.userRecordLoader.Wait
    'Log.out("Loader: end")
    Me.userRecordLoader.Quit
  End Sub
  
  ''' <summary>
  ''' 処理をキャンセルする。
  ''' </summary>
  Public Sub Cancel()
    Me.userRecordLoader.Cancel
  End Sub
  
End Class