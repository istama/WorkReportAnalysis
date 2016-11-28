'
' 日付: 2016/11/25
'
Imports System.IO

''' <summary>
''' SaveFileDialogのヘルパークラス。
''' </summary>
Public Class MySaveFileDialog
  Private sfd As New SaveFileDialog
  
  Public Sub New()
    'はじめに表示されるフォルダを指定する
    '指定しない（空の文字列）の時は、現在のディレクトリが表示される
    sfd.InitialDirectory = "C:\Users\" & System.Environment.UserName & "\Desktop"
    '[ファイルの種類]に表示される選択肢を指定する
    sfd.Filter = "CSVファイル(*.csv)|*.csv|Textファイル(*.txt)|*.txt|すべてのファイル(*.*)|*.*"
    '[ファイルの種類]ではじめに選択されるものを指定する
    '2番目の「すべてのファイル」が選択されているようにする
    sfd.FilterIndex = 1
    'タイトルを設定する
    sfd.Title = "保存先のファイルを選択してください"
    'ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
    sfd.RestoreDirectory = True
    '既に存在するファイル名を指定したとき警告する
    'デフォルトでTrueなので指定する必要はない
    sfd.OverwritePrompt = True
    '存在しないパスが指定されたとき警告を表示する
    'デフォルトでTrueなので指定する必要はない
    sfd.CheckPathExists = True
  End Sub
  
  ''' <summary>
  ''' ファイル保存ダイアログを開き、保存するファイルに出力するためのストリームを返す。
  ''' </summary>
  Public Function Save(fileName As String) As SaveFileStream
    'はじめのファイル名を指定する
    'はじめに「ファイル名」で表示される文字列を指定する
    sfd.FileName = fileName
    
    'ダイアログを表示する
    If sfd.ShowDialog() = DialogResult.OK Then
      'OKボタンがクリックされたとき、
      '選択された名前で新しいファイルを作成し、
      '読み書きアクセス許可でそのファイルを開く。
      '既存のファイルが選択されたときはデータが消える恐れあり。
      Dim stream As System.IO.Stream = sfd.OpenFile
      If Not (stream Is Nothing) Then
        Return New SaveFileStream(stream)
      End If      
    End If
    
    Return Nothing
  End Function
End Class

''' <summary>
''' ファイルに出力するためのストリームクラス。
''' </summary>
Public Class SaveFileStream
  Private stream As Stream
  Private writer As StreamWriter
  
  Private opened As Boolean
  
  Public Sub New(stream As Stream)
    If stream Is Nothing Then Throw New ArgumentNullException("stream is null")
    
    Me.stream = stream
    Me.opened = False
  End Sub
  
  ''' <summary>
  ''' ファイルストリームを開く。
  ''' 書き込む前に必ず呼び出す必要がある。
  ''' </summary>
  Public Sub Open()
    If Not Me.opened Then
      Me.writer = New StreamWriter(stream, System.Text.Encoding.GetEncoding("shift_jis"))
      Me.opened = True
    End If
  End Sub
  
  ''' <summary>
  ''' ファイルストリームを閉じる。
  ''' 終了時に必ず呼び出す必要がある。
  ''' </summary>
  Public Sub Close()
    If Me.opened Then
      Me.writer.Close()
      Me.stream.Close()
      Me.opened = False
    End If
  End Sub
  
  ''' <summary>
  ''' テキストをファイルに追記する。
  ''' </summary>
  Public Sub Write(text As String)
    If Not Me.opened Then Throw New InvalidOperationException("ストリームがまだ開かれていません。")
    
    Me.writer.writeLine(text)
  End Sub
End Class