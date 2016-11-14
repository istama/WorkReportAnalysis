''
'' 日付: 2016/06/25
''
'Imports NUnit.Framework
'Imports NSubstitute
'
'Imports System.Reflection
'Imports System.ComponentModel
'Imports System.Data
'Imports Common.Account
'Imports Common.COM
'Imports Common.IO
'
'<TestFixture> _
'Public Class TestLoader
'	Private filepath As String = ".\test\testForTestLoader.properties"	
'	Private mock As IExcel = Substitute.For(Of IExcel)()
'	Private reader As UserRecordReader = CreateUserRecordReader(CreateMyProperty())
'	Private users As List(Of UserInfo) = CreateUserInfoList()
'		
'	Private Function CreateMyProperty() As MyProperties
'		Dim p As New Properties(filepath)
'		p.Add(MyProperties.KEY_EXCEL_FILEDIR, ".")
'		p.Add(MyProperties.KEY_EXCEL_FILENAME, "{0}.xls")
'		p.Add(MyProperties.KEY_EXCEL_SHEETNAME, "month{0}")
'		p.Add(MyProperties.KEY_ROW_OF_FIRSTDAY_IN_A_MONTH, "1")
'		p.Add(MyProperties.KEY_SELECTABLE_MIN_DATE, "20161001")
'		p.Add(MyProperties.KEY_SELECTABLE_MAX_DATE, "20161130")
'		p.Add(MyProperties.KEY_WORKDAY_COL,                    "Z")
'		p.Add(MyProperties.KEY_ITEM_NAME & "1",                "work1")
'		p.Add(MyProperties.KEY_WORKCOUNT_COL_OF_ITEM & "1",    "A")
'		p.Add(MyProperties.KEY_WORKTIME_COL_OF_ITEM & "1",     "B")
'		p.Add(MyProperties.KEY_PRODUCTIVITY_COL_OF_ITEM & "1", "C")		
'		p.Add(MyProperties.KEY_ITEM_NAME & "2",                "work2")
'		p.Add(MyProperties.KEY_WORKCOUNT_COL_OF_ITEM & "2",    "D")
'		p.Add(MyProperties.KEY_WORKTIME_COL_OF_ITEM & "2",     "E")
'		p.Add(MyProperties.KEY_PRODUCTIVITY_COL_OF_ITEM & "2", "F")	
'		
'		Return New MyProperties(filepath)
'	End Function
'		
'	Private Function CreateUserRecordReader(p As MyProperties) As UserRecordReader
'		Dim rr As New UserRecordReader(p)
'		Dim reader As New ExcelReader(mock)
'		
'		' リフレクションでExcelReaderのモックをプライベートフィールドにセット
'		Dim t As Type = GetType(UserRecordReader)
'		Dim field As FieldInfo = t.GetField(
'			"reader",
'			BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance
'			)
'		field.SetValue(rr, reader)
'		
'		Return rr		
'	End Function
'	
'	Private Function CreateUserInfoList() As List(Of UserInfo)
'		Dim l As New List(Of UserInfo)
'		
'		l.Add(New UserInfo("yamada",    "001", "path"))
'		l.Add(New UserInfo("suzuki",    "002", "path"))
'		l.Add(New UserInfo("sato",      "003", "path"))
'		l.Add(New UserInfo("takahashi", "004", "path"))
'		l.Add(New UserInfo("kimura",    "005", "path"))
'		l.Add(New UserInfo("saito",     "006", "path"))
'		l.Add(New UserInfo("ueno",      "007", "path"))
'		l.Add(New UserInfo("watanabe",  "008", "path"))
'		l.Add(New UserInfo("tanaka",    "009", "path"))
'		l.Add(New UserInfo("ito",       "010", "path"))
'		l.Add(New UserInfo("kobayashi", "011", "path"))
'		l.Add(New UserInfo("kato",      "012", "path"))
'		l.Add(New UserInfo("yoshida",   "013", "path"))
'		l.Add(New UserInfo("shimizu",   "014", "path"))
'		
'		Return l
'	End Function
'	
'	<Test> _
'	Public Sub TestLoad
''		Dim rr As UserRecordReader = CreateUserRecordReader(CreateMyProperty())
''		Dim ul As List(Of UserInfo) = CreateUserInfoList()
'		
'		' Excelクラスのモック実装
'		For Each u In users
'			Dim fileName As String = String.Format(".\{0}.xls", u.GetId)
'			Dim tmp As Integer = Integer.Parse(u.GetId) + 1
'			For m = 10 To 10
'				Dim sheetName As String = String.Format("month{0}", m)
'				For row = 1 To 31
'					If row Mod tmp = 0 Then
'						mock.Read(fileName, sheetName, Cell.Create(row, "Z")).Returns("")
'					Else
'						mock.Read(fileName, sheetName, Cell.Create(row, "Z")).Returns("○")
'					End If 
'					Dim cnt1 As Integer = 60 + tmp - row
'					mock.Read(fileName, sheetName, Cell.Create(row, "A")).Returns(cnt1.ToString)
'					mock.Read(fileName, sheetName, Cell.Create(row, "B")).Returns("6")
'					mock.Read(fileName, sheetName, Cell.Create(row, "C")).Returns(Math.Round((cnt1 / 6), 2).ToString)
'					Dim cnt2 As Integer = 40 - tmp + row
'					mock.Read(fileName, sheetName, Cell.Create(row, "D")).Returns(cnt2.ToString)
'					mock.Read(fileName, sheetName, Cell.Create(row, "E")).Returns("2")
'					mock.Read(fileName, sheetName, Cell.Create(row, "F")).Returns(Math.Round((cnt2 / 2), 2).ToString)
'				Next
'			Next
'		Next
'		
'		Dim loader As New Loader(reader, users)
'		
'		Dim start = Stopwatch.StartNew
'		loader.Load2(New TObserver())
'		start.Stop
'		Console.WriteLine(start.Elapsed.ToString)	
'		
'		Dim urm As UserRecordManager = loader.UserRecordManager
'		' workday = 行 Mod (id + 1)
'		' work1 = 60 + id + 1 - 行, work2 = 40 - (id + 1) + 行
'		Dim r1 As UserRecord = urm.GetUserRecord("001")
'		Assert.AreEqual("61", r1.GetTable("month10").Rows(0)("work1" & vbCrLf & "件数"))
'		Assert.AreEqual("39", r1.GetTable("month10").Rows(0)("work2" & vbCrLf & "件数"))
'		Assert.AreEqual("",   r1.GetTable("month10").Rows(1)("work1" & vbCrLf & "件数"))
'		Assert.AreEqual("41", r1.GetTable("month10").Rows(2)("work2" & vbCrLf & "件数"))
'		Dim r2 As UserRecord = urm.GetUserRecord("002")
'		Assert.AreEqual("62", r2.GetTable("month10").Rows(0)("work1" & vbCrLf & "件数"))
'		Assert.AreEqual("38", r2.GetTable("month10").Rows(0)("work2" & vbCrLf & "件数"))
'		Assert.AreEqual("61", r2.GetTable("month10").Rows(1)("work1" & vbCrLf & "件数"))
'		Assert.AreEqual("",   r2.GetTable("month10").Rows(2)("work2" & vbCrLf & "件数"))
'		Dim r3 As UserRecord = urm.GetUserRecord("003")
'		Assert.AreEqual("",   r3.GetTable("month10").Rows(3)("work1" & vbCrLf & "件数"))
'		Assert.AreEqual("41", r3.GetTable("month10").Rows(4)("work2" & vbCrLf & "件数"))
'		Dim r4 As UserRecord = urm.GetUserRecord("004")
'		Assert.AreEqual("59", r4.GetTable("month10").Rows(5)("work1" & vbCrLf & "件数"))
'		Assert.AreEqual("42", r4.GetTable("month10").Rows(6)("work2" & vbCrLf & "件数"))		
'		Dim r5 As UserRecord = urm.GetUserRecord("005")
'		Assert.AreEqual("58", r5.GetTable("month10").Rows(7)("work1" & vbCrLf & "件数"))
'		Assert.AreEqual("43", r5.GetTable("month10").Rows(8)("work2" & vbCrLf & "件数"))	
'		Dim r6 As UserRecord = urm.GetUserRecord("006")
'		Assert.AreEqual("57", r6.GetTable("month10").Rows(9)("work1" & vbCrLf & "件数"))
'		Assert.AreEqual("44", r6.GetTable("month10").Rows(10)("work2" & vbCrLf & "件数"))
'		Dim r7 As UserRecord = urm.GetUserRecord("007")
'		Assert.AreEqual("56", r7.GetTable("month10").Rows(11)("work1" & vbCrLf & "件数"))
'		Assert.AreEqual("45", r7.GetTable("month10").Rows(12)("work2" & vbCrLf & "件数"))		
'		Dim r8 As UserRecord = urm.GetUserRecord("008")
'		Assert.AreEqual("55", r8.GetTable("month10").Rows(13)("work1" & vbCrLf & "件数"))
'		Assert.AreEqual("46", r8.GetTable("month10").Rows(14)("work2" & vbCrLf & "件数"))
'		Dim r9 As UserRecord = urm.GetUserRecord("009")
'		Assert.AreEqual("54", r9.GetTable("month10").Rows(15)("work1" & vbCrLf & "件数"))
'		Assert.AreEqual("47", r9.GetTable("month10").Rows(16)("work2" & vbCrLf & "件数"))
'		Dim r10 As UserRecord = urm.GetUserRecord("010")
'		Assert.AreEqual("53", r10.GetTable("month10").Rows(17)("work1" & vbCrLf & "件数"))
'		Assert.AreEqual("48", r10.GetTable("month10").Rows(18)("work2" & vbCrLf & "件数"))		
'		Dim r11 As UserRecord = urm.GetUserRecord("011")
'		Assert.AreEqual("52", r11.GetTable("month10").Rows(19)("work1" & vbCrLf & "件数"))
'		Assert.AreEqual("49", r11.GetTable("month10").Rows(20)("work2" & vbCrLf & "件数"))	
'		Dim r12 As UserRecord = urm.GetUserRecord("012")
'		Assert.AreEqual("51", r12.GetTable("month10").Rows(21)("work1" & vbCrLf & "件数"))
'		Assert.AreEqual("50", r12.GetTable("month10").Rows(22)("work2" & vbCrLf & "件数"))
'		Dim r13 As UserRecord = urm.GetUserRecord("013")
'		Assert.AreEqual("50", r13.GetTable("month10").Rows(23)("work1" & vbCrLf & "件数"))
'		Assert.AreEqual("51", r13.GetTable("month10").Rows(24)("work2" & vbCrLf & "件数"))		
'		Dim r14 As UserRecord = urm.GetUserRecord("014")
'		Assert.AreEqual("49", r14.GetTable("month10").Rows(25)("work1" & vbCrLf & "件数"))
'		Assert.AreEqual("52", r14.GetTable("month10").Rows(26)("work2" & vbCrLf & "件数"))	
'		Assert.AreEqual("47", r14.GetTable("month10").Rows(27)("work1" & vbCrLf & "件数"))
'		Assert.AreEqual("54", r14.GetTable("month10").Rows(28)("work2" & vbCrLf & "件数"))	
'		Assert.AreEqual("",   r14.GetTable("month10").Rows(29)("work1" & vbCrLf & "件数"))	
'		Assert.AreEqual("56", r14.GetTable("month10").Rows(30)("work2" & vbCrLf & "件数"))
'	End Sub
'	
'	<Test> _
'	Public Sub TestThreadObserver
'		Dim loader As New Loader(reader, users)
'		Dim observer As New TObserver()
'		
'		loader.Load(observer)
'		
'		Assert.AreEqual(users.Count, observer.Count)
'	End Sub
'	
'	<Test> _
'	Public Sub TestLoadCancel
'		Dim loader As New Loader(reader, users)
'		Dim observer As New TObserver()
'		
'		observer.CancelAsync
'		loader.Load(observer)
'		
'		Assert.AreEqual(0, observer.Count)
'	End Sub
'	
'	<TestFixtureTearDownAttribute>
'	Public Sub TearDown
'		DeleteTestFile(filepath)
'	End Sub
'	
'	Public Sub DeleteTestFile(filepath As String)
'		If System.IO.File.Exists(filepath) Then
'			System.IO.File.Delete(filepath)
'		End If
'	End Sub
'	
'	Private Class TObserver
'		Implements IThreadObserver
'		
'		Public Count As Integer
'		Public Canceled As Boolean
'		
'		Public Sub New()
'			Count = 0
'			Canceled = False			
'		End Sub
'		
'		Public Sub ReportProgress(args As Object) Implements IThreadObserver.ReportProgress
'			SyncLock Me
'				Count += 1
'			End SyncLock
'		End Sub
'		
'		Public Function CancellationPending() As Boolean Implements IThreadObserver.CancellationPending
'			Return Canceled
'		End Function
'		
'		Public Sub CancelAsync() Implements IThreadObserver.CancelAsync
'			Canceled = True
'		End Sub
'	End Class
'End Class
