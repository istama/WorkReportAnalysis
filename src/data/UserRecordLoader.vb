'
' 日付: 2016/10/18
'
Imports System.Threading.Tasks
Imports System.Collections.Concurrent

Imports Common.Account
Imports Common.COM
Imports Common.IO

''' <summary>
''' ユーザレコードを読み込んでバッファに格納するクラス。
''' </summary>
Public Class UserRecordLoader
  Private ReadOnly excel As Excel4
  
  ''' 読み込んだユーザレコードを格納しておくクラス
  Private ReadOnly userRecordBuffer As UserRecordBuffer
  ''' ユーザレコードにアクセスするクラス
  Private ReadOnly userRecordReader As UserRecordReader
  
  ''' <summary>
  ''' ユーザレコードを読み込むタスクを保持するキュー。
  ''' （現在、マルチスレッドで実装されていないので使用していない。）
  ''' </summary>
  Private ReadOnly tasks As New ConcurrentQueue(Of Task)
  
  Public Sub New(properties As ExcelProperties)
    If properties Is Nothing Then Throw New ArgumentNullException("properties is null")
    
    Me.excel = New Excel4()
    Me.userRecordBuffer = New UserRecordBuffer(properties)
    Me.userRecordReader = New UserRecordReader(properties, Me.excel)
  End Sub
  
  ''' <summary>
  ''' 読み込んだユーザレコードを保持するオブジェクトを返す。
  ''' </summary>
  Public Function GetUserRecordBuffer() As UserRecordBuffer
    Return Me.userRecordBuffer
  End Function
  
  ''' <summary>
  ''' 初期処理を行う。
  ''' ユーザレコードを読み込む前に必ず実行すること。
  ''' </summary>
  Public Sub Init()
    #If Debug = False Then
      Me.excel.init
    #End If
  End Sub
  
  ''' <summary>
  ''' 終了処理を行う。
  ''' ユーザレコードの読み込み終了時には必ず実行すること。
  ''' </summary>
  Public Sub Quit()
    #If Debug = False Then
      Me.excel.Quit
    #End If
  End Sub
  
  ''' <summary>
  ''' ユーザレコードがすべて読み込まれるまで待機する。
  ''' </summary>
  Public Sub Wait
    While True
      Threading.Thread.Sleep(50)
      ' 現在マルチスレッドで実装されていないので使用されない
      If Task.WaitAll(Me.tasks.ToArray, 0) Then
        Exit While
      End If
    End While
  End Sub
  
  ''' <summary>
  ''' 指定したユーザのレコードを読み込む。
  ''' 読み込んだレコードはUserRecordBufferに格納される。
  ''' </summary>
  Public Sub Load(loadedUserInfo As UserInfo)
    Load(loadedUserInfo, Nothing)
  End Sub
  
  ''' <summary>
  ''' 指定したユーザのレコードを読み込む。
  ''' 読み込んだレコードはUserRecordBufferに格納される。
  ''' 読み込みが終了するとcallback関数が呼び出される。
  ''' </summary>
  Public Sub Load(loadedUserInfo As UserInfo, callback As Action(Of UserRecord))
    'tasks.Enqueue(Task.Factory.StartNew(Sub() LoadTask(loadedUserInfo, f)))
    LoadTask(loadedUserInfo, callback)
  End Sub
  
  ''' <summary>
  ''' ユーザレコードの読み込みを実際に行うメソッド。
  ''' </summary>
  Private Sub LoadTask(loadedUserInfo As UserInfo, callback As Action(Of UserRecord))
    Try
      ' ユーザ情報オブジェクトからユーザレコードオブジェクトを取得する
      ' この生成したオブジェクトに読み込んだレコードをセットする
      Dim record As UserRecord = Me.userRecordBuffer.GetUserRecord(loadedUserInfo)
      ' 全ユーザレコードの集計値から現在のレコードの値を減算する
      ' (再読み込みの場合、集計値を計算しなおすために減算している）
      Me.userRecordBuffer.MinusToTotalRecord(loadedUserInfo)
      ' ユーザレコードを読み込む
      Me.userRecordReader.Read(record)
      ' 全ユーザレコードの集計値に読み込んだレコードの値を加算する
      Me.userRecordBuffer.PlusToTotalRecord(loadedUserInfo)
      
      If callback IsNot Nothing Then
        callback(record)
      End If
    Catch ex As Exception
      Log.out(ex.Message)
    End Try
  End Sub
  
  ''' <summary>
  ''' 読み込みをキャンセルする。
  ''' </summary>
  Public Sub Cancel()
    Me.userRecordReader.Cancel
  End Sub
End Class
