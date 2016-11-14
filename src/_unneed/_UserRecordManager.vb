''
'' 日付: 2016/06/22
''
'Public Class UserRecordManager
'	Private _userDataDictionary As IDictionary(Of String, UserData)
'	Public ReadOnly Property UserDataList As List(Of UserData)
'		Get
'			Return New List(Of UserData)(_userDataDictionary.Values)
'		End Get
'	End Property
'	
'	Public Sub New()
'		Me._userDataDictionary = New Dictionary(Of String, UserData)
'	End Sub
'	
'	Public Sub Add(name As String, id As String, record As UserRecord)
'		If name Is Nothing Then
'			Throw New ArgumentNullException("name is null")
'		End If
'		If id Is Nothing Then
'			Throw New ArgumentNullException("id is null")
'		End If
'		If record Is Nothing Then
'			Throw New ArgumentNullException("record Is null")
'		End If
'		
'		Dim data As New UserData(name, id, record)
'		_userDataDictionary.Add(id, data)
'	End Sub
'	
'	Public Function GetUserRecord(id As String) As UserRecord
'		If id Is Nothing Then
'			Throw New ArgumentNullException("id is null")
'		End If
'		
'		Dim data As UserData = Nothing
'		If Not _userDataDictionary.TryGetValue(id, data) Then
'			Throw New KeyNotFoundException("id dose not exists / " & id)
'		End If
'		
'		Return data.Record
'	End Function
'End Class
'
'Public Class UserData
'	Private _Name As String
'	Public ReadOnly Property Name As String
'		Get
'			Return _Name
'		End Get
'	End Property
'	
'	Private _Id As String
'	Public ReadOnly Property Id As String
'		Get
'			Return _Id
'		End Get
'	End Property
'	
'	Private _Record As UserRecord
'	Public ReadOnly Property Record As UserRecord
'		Get
'			Return _Record
'		End Get
'	End Property
'	
'	Sub New(name As String, id As String, record As UserRecord)
'		Me._Name   = name
'		Me._Id     = id
'		Me._Record = record
'	End Sub
'End Class