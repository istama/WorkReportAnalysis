'
' 日付: 2016/10/18
' 
Imports System.ComponentModel
Imports Common.IO

''' <summary>
''' ファイルのロード状況を監視し、callbackに通知するクラス。
''' </summary>
Public Class ThreadObserver
  Private ReadOnly control As Control  
  Private ReadOnly callback As Action(Of Integer, String)
  
  ' 処理の進捗状況を表す値
	Private barMeter As Integer
	
	Public Sub New(control As Control, callback As Action(Of Integer, String))
	  If control Is Nothing Then Throw New ArgumentNullException("control is null")
	  If callback Is Nothing Then Throw New ArgumentNullException("callback is null")
	  
	  Me.control = Control
	  Me.callback = callback
	  Me.barMeter = 0
	End Sub
	
	''' <summary>
	''' 進捗状況をワーカーに伝える。
	''' </summary>
	Public Sub ReportProgress(msg As String)
		SyncLock Me
		  barMeter += 1
		  'Log.out("ReportProgress: " & msg & "  " & barMeter.ToString)
		  Notification(barMeter, msg)
		End SyncLock
	End Sub
	
	Private Sub Notification(meter As Integer, msg As String)
	  If Me.control.InvokeRequired Then
	    Me.control.Invoke(Sub() Notification(meter, msg))
	  Else
	    Me.callback(meter, msg)
	  End If
	End Sub
	
End Class
