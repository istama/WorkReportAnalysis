'
' 日付: 2016/06/20
'
Imports System.ComponentModel

''' <summary>
''' 全ユーザのExcelファイルが読み込まれる時に表示されるプログレスバー。
''' </summary>
Public Partial Class ProgressBarForm
  ''' バックグラウンドで行うタスクをもつオブジェクト  
	Private _loader As Loader
	Public Property Loader() As Loader
	  Get
	    Return _loader
	  End Get
	  
		Set(loader As Loader)
			_loader = loader
		End Set
	End Property
	
	Private StopWatch As System.Diagnostics.Stopwatch
		
	Public Sub New()
		Me.InitializeComponent()
	End Sub
	
	Sub ProgressBarFormLoad(sender As Object, e As EventArgs)
		'処理が行われているときは、何もしない
		If BackgroundWorker1.IsBusy Then Return
		
		If _loader Is Nothing Then Return

    'コントロールを初期化する
    Me.pBar.Minimum = 0
    Me.pBar.Maximum = _loader.LoadedCount
    Me.pBar.Value   = 0

    'BackgroundWorkerのProgressChangedイベントが発生するようにする
    BackgroundWorker1.WorkerReportsProgress = True
    
    ' 読み込みにかかった時間を計測する
    Stopwatch = System.Diagnostics.Stopwatch.StartNew()
    
    'DoWorkで取得できるパラメータを指定して処理を開始する
    'パラメータが必要なければ省略できる
    BackgroundWorker1.RunWorkerAsync(_loader)
	End Sub
	
  ''' <summary>
  ''' BackgroundWorker1のDoWorkイベントハンドラ。
  ''' ここで時間のかかる処理を行う。
  ''' </summary>
  ''' <param name="sender"></param>
  ''' <param name="e"></param>
  Sub BackgroundWorker1DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs)
  	Dim bgWorker As BackgroundWorker = DirectCast(sender, BackgroundWorker)
  	' RunWorkerAsync()で渡したパラメータを取得する
    Dim loader As Loader = DirectCast(e.Argument, Loader)
    
    Try
      loader.Load(
        New ThreadObserver(
          Me,
          Sub(meter, msg)
            Me.pBar.Value = meter
            Me.lblMsg.Text = msg
          End Sub))
    Catch ex As Exception
    	MsgBox.ShowError(ex)
    Finally
      loader.Quit
    End Try
	End Sub
  
  'BackgroundWorker1のProgressChangedイベントハンドラ
  'コントロールの操作は必ずここで行い、DoWorkでは絶対にしない
  ' ※ＵＩの更新が綺麗にいかないのでこのProgressChangedの機構は使わない。
  '   かわりにInvoke()を使用する方法を自前で用意。
'  Private Sub BackgroundWorker1ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs)
'    'ProgressBarの値を変更する
'    Me.pBar.Value = e.ProgressPercentage
'    ' 進捗状況をラベルに表示する
'    'ChangeLabel(DirectCast(e.UserState, String))
'    Me.lblMsg.Text = DirectCast(e.UserState, String)
'  End Sub
  
  'BackgroundWorker1のRunWorkerCompletedイベントハンドラ
  '処理が終わったときに呼び出される
  Private Sub BackgroundWorker1RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs)
    Stopwatch.Stop()
    'MessageBox.Show("complete")

    If Not e.Error Is Nothing Then
      'エラーが発生したとき
      'Label1.Text = "エラー:" & e.Error.Message
    Else
      '正常に終了したとき
      '結果を取得する
      'MessageBox.Show("読み込み時間: " & Stopwatch.Elapsed.ToString)

      'Dim result As Integer = CInt(e.Result)
      'Label1.Text = result.ToString() & "回で完了しました。"
    End If

    Me.Close()
  End Sub

	
	Sub BtnCancelClick(sender As Object, e As EventArgs)
		Me._loader.Cancel	
	End Sub
	
	Private Delegate Sub SetStringDelegate(str As String)
	
End Class
