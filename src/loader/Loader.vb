'
' 日付: 2016/10/18
'
Imports Common.Account
Imports Common.IO

Public Class Loader
  Private userInfoManager As UserInfoManager
  Private userRecordManager As UserRecordManager
  Private userRecordLoader As UserRecordLoader
  
  Private _cancel As Boolean = False
  
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
    Return Me.UserInfoManager.UserInfoCount
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
    For Each ui As UserInfo In Me.userInfoManager.UserInfos
      If Me._cancel Then
        Return
      End If
      
      ' 指定したユーザの情報を読み込み、読み込んだユーザの情報をオブザーバーに通知する
      Me.userRecordLoader.Load(
        ui,
        Sub(record) observer.ReportProgress(record.GetIdNumber & " " & record.GetName))     
    Next
    
    Me.userRecordLoader.Wait
    Me.userRecordLoader.Quit
  End Sub
  
  ''' <summary>
  ''' 処理をキャンセルする。
  ''' </summary>
  Public Sub Cancel()
    Me.userRecordLoader.Cancel
    Me._cancel = True
  End Sub
  
End Class
