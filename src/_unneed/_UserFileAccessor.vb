''
'' 日付: 2016/06/20
''
'Imports Common.Account
'
'Public Class UserFileAccessor
'	Private properties As MyProperties
'	Private loader As Loader
'	
'	Public Sub New(properties As MyProperties)
'		Me.properties = properties
'	End Sub
'	
'	''' <summary>
'	''' 全てのユーザ情報を読み込む。
'	''' </summary>
'	Public Function ReadAllUserInfo() As List(Of UserInfo)
'		Return UserInfoManager.Create(properties.UserFilePath).UserInfoList
'	End Function
'	
'	''' <summary>
'	''' 全てのユーザのExcelファイルを読み込む。
'	''' </summary>
'	Public Function ReadAllUserRecord(userInfoList As List(Of UserInfo)) As UserRecordManager
'		If userInfoList Is Nothing Then
'			Throw New ArgumentNullException("userInfoList is null")
'		End If
'		
'		If loader Is Nothing Then
'			' Excelファイルのローダーを生成
'			Dim reader As New UserRecordReader(properties)
'			loader = New Loader(reader, userInfoList)
'		End If
'		
'		Dim res As DialogResult =
'			MessageBox.Show("全てのExcelファイルを読み込みますか？" & vbCrLf & "読み込みには時間がかかるかもしれません。", "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)
'    If res = DialogResult.OK Then
'    	ProgressBarForm.Loader = loader
'      ProgressBarForm.ShowDialog()
'    End If
'    
'    Return loader.UserRecordManager
'	End Function
'End Class
