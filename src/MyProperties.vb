'
' 日付: 2016/05/05
'
Imports Common.IO
Imports Common.Util

Public Class MyProperties
  Inherits AppProperties
  
  ' プロパティのキー値（Constを付けたフィールドは静的定数になる）
  Public Const KEY_APPLICATION_TITLE    = "ApplicationTitle"
  Public Const KEY_USERS_FILEPATH       = "UsersFilePath"
  
  Public Const KEY_VERSION_FILEPATH     = "VersionFilePath"
  Public Const KEY_LATEST_EXE_FILES_DIR = "LatestExeFilesDir"
  Public Const KEY_IS_UPDATE_RUNNABLE   = "IsAutoUpdateRunnable"
  
  ''' <summary>
  ''' コンストラクタ。
  ''' </summary>
  ''' <param name="filePath">このアプリケーションのプロパティファイルのパス</param>
  Public Sub New(filePath As String)
    MyBase.New(filePath)
  End Sub
  
  ''' <summary>
  ''' プロパティのデフォルト値。
  ''' </summary>
  ''' <returns></returns>
  Protected Overrides Function DefaultProperties() As IDictionary(Of String, String)
    Dim p As New Dictionary(Of String, String)
    
    p.Add(KEY_APPLICATION_TITLE, "WorkReport")				
    p.Add(KEY_USERS_FILEPATH, ".\userinfo.txt")
    
    p.Add(KEY_IS_UPDATE_RUNNABLE, "False")
    p.Add(KEY_VERSION_FILEPATH, "")
    p.Add(KEY_LATEST_EXE_FILES_DIR, "")
    
    Return p
  End Function
  
  ''' <summary>
  ''' デフォルトにないプロパティがプロパティファイルにあることを認めるかどうか。
  ''' </summary>
  ''' <returns></returns>
  Protected Overrides Function AllowNonDefaultProperty() As Boolean
    Return False
  End Function
  
  Public Function UserFilePath As String
    Return Me.GetOrDefault(KEY_USERS_FILEPATH, String.Empty)
  End Function
  
End Class
