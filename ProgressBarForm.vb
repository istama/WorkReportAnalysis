
Imports System.ComponentModel
Imports MP.WorkReportAnalysis.App
Imports MP.WorkReportAnalysis.Model
Imports MP.WorkReportAnalysis.Control

Public Class ProgressBarForm
  Private _UserRecordLoader As UserRecordLoader
  Public WriteOnly Property UserRecordLoader() As UserRecordLoader
    Set(value As UserRecordLoader)
      _UserRecordLoader = value
    End Set
  End Property

  Private Stopwatch As System.Diagnostics.Stopwatch

  Private Sub ProgressBarForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    '処理が行われているときは、何もしない
    If BackgroundWorker1.IsBusy Then
      Return
    End If

    'コントロールを初期化する
    ProgressBar1.Minimum = 0
    ProgressBar1.Maximum = _UserRecordLoader.UserRecordManager.GetUserInfoList.Count
    ProgressBar1.Value = 0

    'BackgroundWorkerのProgressChangedイベントが発生するようにする
    BackgroundWorker1.WorkerReportsProgress = True

    'DoWorkで取得できるパラメータを指定して処理を開始する
    'パラメータが必要なければ省略できる
    Stopwatch = System.Diagnostics.Stopwatch.StartNew()
    BackgroundWorker1.RunWorkerAsync(_UserRecordLoader)
  End Sub

  'BackgroundWorker1のDoWorkイベントハンドラ
  'ここで時間のかかる処理を行う
  Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork
    Dim bgWorker As BackgroundWorker = DirectCast(sender, BackgroundWorker)
    Dim loader As UserRecordLoader = DirectCast(e.Argument, UserRecordLoader)

    '時間のかかる処理を開始する
    Dim meter As Integer = 1
    For Each user As ExpandedUserInfo In loader.UserRecordManager.GetUserInfoList
      If bgWorker.CancellationPending Then
        e.Cancel = True
        Return
      Else
        Try
          loader.Load(user)
          'System.Threading.Thread.Sleep(1000)
        Catch ex As Exception
          Dim res As DialogResult = MessageBox.Show(ex.Message & vbCrLf & vbCrLf & "ファイルの読み込みを続けますか？", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
          If res = DialogResult.No Then
            Exit For
          End If
        End Try

        'ProgressChangedイベントハンドラを呼び出し、
        'コントロールの表示を変更する
        bgWorker.ReportProgress(meter)
        meter += 1
      End If
    Next

  End Sub

  'BackgroundWorker1のProgressChangedイベントハンドラ
  'コントロールの操作は必ずここで行い、DoWorkでは絶対にしない
  Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
    'ProgressBar1の値を変更する
    Dim value As Integer = e.ProgressPercentage
    ProgressBar1.Value = value
    'Dim users As List(Of ExpandedUserInfo) = _UserRecordManager.GetUserInfoList()
    'lblFileName.Text = users(value).GetIdNum
  End Sub

  'BackgroundWorker1のRunWorkerCompletedイベントハンドラ
  '処理が終わったときに呼び出される
  Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
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

  Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
    BackgroundWorker1.CancelAsync()
  End Sub

End Class