''
'' 日付: 2016/06/21
''
'Imports System.ComponentModel
'Imports System.Data
'Imports System.Threading
'Imports System.Threading.Tasks
'Imports Common.Account
'
'Public Class Loader
'	Implements ILoader
'	Private userRecordReader As UserRecordReader
'	Private userInfoList As List(Of UserInfo)
'
'	Private _userRecordManager As UserRecordManager
'	Public ReadOnly Property UserRecordManager As UserRecordManager
'		Get
'			Return _userRecordManager
'		End Get
'	End Property
'	
'	Public Sub New(userRecordReader As UserRecordReader, userInfoList As List(Of UserInfo))
'		Me._userRecordManager = New UserRecordManager()
'		Me.userRecordReader = userRecordReader
'		
'		Me.userInfoList = New List(Of UserInfo)		
'		For Each ui In userInfoLIst
'			Me.userInfoList.Add(ui)
'		Next
'	End Sub
'	
'	''' <summary>
'	''' 読み込むファイルの件数。
'	''' </summary>
'	''' <returns></returns>
'	Public Function LoadedCount() As Integer Implements ILoader.LoadedCount
'		Return userInfoList.Count
'	End Function
'	
'	''' <summary>
'	''' 全ユーザのファイルを読み込み、UserRecordManagerに登録する。
'	''' 引数のthreadObseverから、ファイルを１件読み込むたびに通知を受けることができる。
'	''' </summary>
'	''' <param name="threadObserver">通知用オブジェクト</param>
'	Public Sub Load(threadObserver As IThreadObserver) Implements ILoader.Load
'		'userRecordReader.Init()
'		My.Application.Log.WriteEntry("in loader.load, complete init reader")
'		
'		Dim taskArray As New List(Of Task)
'		' ユーザごとのExcelファイルを読み込むタスクをスレッドプールにセットする
'		For i = 0 To userInfoList.Count - 1
'			Dim args As New Arguments(userInfoList(i), threadObserver, Nothing, Nothing)
'			' タスクを作成、実行する
'			taskArray.Add(Task.Factory.StartNew(Sub(a As Object) Me.LoadTask(a), args))
'		Next
'		My.Application.Log.WriteEntry("in loader.load, complete to start all task")
'		
'		' 全てのファイルが読み込まれるまでブロック
'		Task.WaitAll(taskArray.ToArray)
'		My.Application.Log.WriteEntry("in loader.load, complete all task")
'		
'		'userRecordReader.Quit()
'		My.Application.Log.WriteEntry("in loader.load, complete to quit reader")
'	End Sub
'	
'	Private Sub LoadTask(ByVal args As Object)
'		Dim arguments As Arguments = CType(args, Arguments)
'		Dim userInfo As UserInfo = arguments.UserInfo
'		Dim observer As IThreadObserver = arguments.Observer
'		
'		Try
'			' キャンセルボタンが押されていないことを確認
'			If Not observer.CancellationPending Then
'				' ファイルを読み込む
'				Dim record As UserRecord = userRecordReader.Read(userInfo.GetSimpleId)
'				My.Application.Log.WriteEntry("Loader.LoadTask() " & userInfo.GetName)
'				' 読み込んだデータをマネージャに登録
'				_userRecordManager.Add(userInfo.GetName, userInfo.GetSimpleId, record)
'				' データを読み込んだことを通知
'				observer.ReportProgress(userInfo)
'			End If
'		Catch ex As Exception
'			Dim res As DialogResult =
'				MessageBox.Show(
'					userInfo.GetSimpleId & " " & userInfo.GetName & vbCrLf & vbCrLf &
'					ex.Message & vbCrLf & vbCrLf & "ファイルの読み込みを続けますか？", "警告",
'					MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
'			If res = DialogResult.No Then
'				observer.CancelAsync			
'			End If
'		End Try
'	End Sub
'	
'	''' <summary>
'	''' タスクに渡すための引数クラス
'	''' </summary>
'	Private Class Arguments
'		Public UserInfo As UserInfo
'		Public Observer As IThreadObserver
'		Public CountDown As CountdownEvent
'		Public CancelToken As CancellationTokenSource
'		
'		Public Sub New(userInfo As UserInfo, observer As IThreadObserver, countDown As CountdownEvent, cancelToken As CancellationTokenSource)
'			Me.UserInfo = userInfo
'			Me.Observer = observer
'			Me.CountDown = countDown
'			Me.CancelToken = cancelToken
'		End Sub
'	End Class
'	
'	'
'	'
'	'
'	'
'	'
'	'
'	'
'	'
'	
'	' スレッドプールを用いたファイル読み込み
'	' ソースを残しておく
'	Public Sub Load2(threadObserver As IThreadObserver) 'Implements ILoader.Load
'		Dim countDown As New CountdownEvent(userInfoList.Count)
'		Dim countDownCancelToken As New CancellationTokenSource()
'		
'		userRecordReader.Init()
'		
'		' 同時に実行するスレッド数を設定
'		Dim THREAD_COUNT = 5
'		Dim IO_COUNT = 1024
'		ThreadPool.SetMaxThreads(THREAD_COUNT, IO_COUNT)
''		ThreadPool.SetMinThreads(4, 512)
'		
'		' ユーザごとのExcelファイルを読み込むタスクをスレッドプールにセットする
'		For Each userInfo In userInfoList
'			Dim args As New Arguments(userInfo, threadObserver, countDown, countDownCancelToken)
'			ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf LoadTask2), args)
'		Next
'		
'		' 全てのファイルが読み込まれるまでブロック
''		countDown.Wait(countDownCancelToken)
'		countDown.Wait()
'		countDownCancelToken.Dispose
'		
''		While True
''			' スレッドプール内の空いているスレッドの数を取得
''			Dim threadCnt As Integer
''			Dim ioCnt As Integer
''			ThreadPool.GetAvailableThreads(threadCnt, ioCnt)
''			If threadCnt = THREAD_COUNT AndAlso userDataList.TrueForAll(Function(ud) ud.Record IsNot Nothing) Then
''				Exit while
''			End If
''			Thread.Sleep(500)
''		End While
'		
'		userRecordReader.Quit()
'	End Sub
'	
'	Private Sub LoadTask2(ByVal args As Object)
'		Dim arguments As Arguments = CType(args, Arguments)
'		Dim userInfo As UserInfo = arguments.UserInfo
'		Dim observer As IThreadObserver = arguments.Observer
''		Dim bgWorker As BackgroundWorker = arguments.BackGroundWorker
'		Dim countDown As CountdownEvent = arguments.CountDown
''		Dim countDonwCancelToken As CancellationTokenSource = arguments.CancelToken
'		
'		' キャンセルボタンが押されていないことを確認
''		If Not bgWorker.CancellationPending Then
'			Try
'				If Not observer.CancellationPending Then
'					' ファイルを読み込む
'					Dim record As UserRecord = userRecordReader.Read(userInfo.GetSimpleId)
'					_userRecordManager.Add(userInfo.GetName, userInfo.GetSimpleId, record)
'					
'					observer.ReportProgress(args)
'				End If
'			Catch ex As Exception
'				Dim res As DialogResult =
'					MessageBox.Show(
'						userInfo.GetSimpleId & " " & userInfo.GetName & vbCrLf & vbCrLf &
'						ex.Message & vbCrLf & vbCrLf & "ファイルの読み込みを続けますか？", "警告",
'						MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
'				If res = DialogResult.No Then
'					observer.CancelAsync			
'				End If
'			Finally
'				' カウントをデクリメントする
'				' 0になるとCountDownEvent.Wait()のブロックが解除される
'				countDown.Signal
'			End Try
''		Else
''			If countDonwCancelToken.IsCancellationRequested = False Then
''				SyncLock Me
''					If countDonwCancelToken.IsCancellationRequested = False Then
''						' CountDownEvent.Wait()のブロックを解除する
''						countDonwCancelToken.Cancel
''					End If
''				End SyncLock
''			End If
''		End If
'	End Sub
'End Class
