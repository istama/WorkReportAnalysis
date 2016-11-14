''
'' 日付: 2016/06/23
''
'Imports NUnit.Framework
'Imports NSubstitute
'
'Imports System.Data
'Imports Common.COM
'Imports Common.IO
'Imports System.Reflection
'
'<TestFixture> _
'Public Class TestUserRecordReader
'	Private filepath As String = "./test/test.properties"
'	Private mock As IExcel = Substitute.For(Of IExcel)()	
'		
'	Private Function CreateUserRecordReader(filepath As String) As UserRecordReader
'		Dim rr = New UserRecordReader(New MyProperties(filepath))
'		Dim reader As New ExcelReader(mock)
'		
'		' リフレクションでExcelReaderのモックをプライベートフィールドにセット
'		Dim t As Type = GetType(UserRecordReader)
''		t.InvokeMember(
''			"reader",
''			BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.SetField,
''			Nothing,
''			rr,
''			New Object() { reader }
''			)
'		
'		Dim field As FieldInfo = t.GetField(
'			"reader",
'			BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance
'			)
'		field.SetValue(rr, reader)
'		
'		Return rr
'	End Function
'	
'	<Test> _
'	Public Sub TestRead
'		' Excelファイルを読み込みDataSetにして返す
'		
'		' Excelファイルの設定
'		Dim path As String = "./test/test2.properties"
'		Dim p As New Properties(path)
'		p.Add(MyProperties.KEY_EXCEL_FILEDIR, ".")
'		p.Add(MyProperties.KEY_EXCEL_FILENAME, "file{0}.xls")
'		p.Add(MyProperties.KEY_EXCEL_SHEETNAME, "{0}月分")
'		p.Add(MyProperties.KEY_ROW_OF_FIRSTDAY_IN_A_MONTH, "1")
'		p.Add(MyProperties.KEY_SELECTABLE_MIN_DATE, "20161001")
'		p.Add(MyProperties.KEY_SELECTABLE_MAX_DATE, "20161031")
'		p.Add(MyProperties.KEY_WORKDAY_COL,                    "Z")
'		p.Add(MyProperties.KEY_ITEM_NAME & "1",                "work1")
'		p.Add(MyProperties.KEY_WORKCOUNT_COL_OF_ITEM & "1",    "A")
'		p.Add(MyProperties.KEY_WORKTIME_COL_OF_ITEM & "1",     "B")
'		p.Add(MyProperties.KEY_PRODUCTIVITY_COL_OF_ITEM & "1", "C")
'		' 2項目目以降は読み込まない
'		p.Add(MyProperties.KEY_ITEM_NAME & "2", "")
'		p.Add(MyProperties.KEY_ITEM_NAME & "3", "")
'		p.Add(MyProperties.KEY_ITEM_NAME & "4", "")
'		p.Add(MyProperties.KEY_ITEM_NAME & "5", "")
'		p.Add(MyProperties.KEY_ITEM_NAME & "6", "")
'		p.Add(MyProperties.KEY_ITEM_NAME & "7", "")
'	
'		Dim rr As UserRecordReader = CreateUserRecordReader(path)
'		
'		' Excelクラスのモックの戻り値の設定
'		mock.Read(".\file001.xls", "10月分", Cell.Create(1,  "Z")).Returns("○")
'		mock.Read(".\file001.xls", "10月分", Cell.Create(1,  "A")).Returns("50")
'		mock.Read(".\file001.xls", "10月分", Cell.Create(1,  "B")).Returns("8")
'		mock.Read(".\file001.xls", "10月分", Cell.Create(1,  "C")).Returns("6.22")
'		mock.Read(".\file001.xls", "10月分", Cell.Create(31, "Z")).Returns("○")
'		mock.Read(".\file001.xls", "10月分", Cell.Create(31, "A")).Returns("60")
'		mock.Read(".\file001.xls", "10月分", Cell.Create(31, "B")).Returns("7")
'		mock.Read(".\file001.xls", "10月分", Cell.Create(31, "C")).Returns("8.57")	
'		
'		Dim r As UserRecord = rr.Read("001")
'		
'		Dim table As DataTable = r.GetTable("10月分")
'		Assert.AreEqual(31, table.Rows.Count)
'		
'		Dim row1 As DataRow = table.Rows(0)
'		Assert.AreEqual("50",   row1("work1" & vbCrLf & "件数"))
'		Assert.AreEqual("8",    row1("work1" & vbCrLf & "作業時間"))
'		Assert.AreEqual("6.22", row1("work1" & vbCrLf & "生産性"))
'		
'		Dim row2 As DataRow = table.Rows(30)
'		Assert.AreEqual("60",   row2("work1" & vbCrLf & "件数"))
'		Assert.AreEqual("7",    row2("work1" & vbCrLf & "作業時間"))
'		Assert.AreEqual("8.57", row2("work1" & vbCrLf & "生産性"))
'		
'		DeleteTestFile(path)
'	End Sub
'	
''	<Test> _
''	Public Sub TestCreateColumnNode
''		' プロパティファイルに設定された列と列名から、Excelにアクセスするための列ノードを作成する
''		Dim pp As New Properties(filepath)
''		
''		' これらの項目はプロパティ値がセットされてないため、列ノードには含まれない
''		pp.Add(MyProperties.KEY_ITEM_NAME & "2", "")
''		pp.Add(MyProperties.KEY_WORKCOUNT_COL_OF_ITEM & "4", "")
''		pp.Add(MyProperties.KEY_WORKTIME_COL_OF_ITEM & "6", "")
''		
''		Dim p As New MyProperties(filepath)
''		Dim n As ColumnNode = UserRecordReader.CreateColumnNode(p)
''		
''		Assert.AreEqual("出勤日", n.Name)
''		Dim idx As Integer = 0
''		For i = 1 To p.GetItemValuesList.Count
''			' 列ノードに含まれるインデックス
''			If i = 1 OrElse i = 3 OrElse i = 5 OrElse i = 7 Then
''				CheckWorkItemColumn(p.GetItemName(i), n.Children(idx))
''				idx += 1
''			End If
''		Next
''		Assert.AreEqual(p.GetValue(MyProperties.KEY_NOTE_NAME), n.Children(n.Children.Count - 1).Name)
''	End Sub
''	
''	Private Sub CheckWorkItemColumn(colName As String, node As ColumnNode)
''		Assert.AreEqual(colName & vbCrLf & "件数",     node.Name)
''		Assert.AreEqual(colName & vbCrLf & "作業時間", node.Children(0).Name)
''		Assert.AreEqual(colName & vbCrLf & "生産性",   node.Children(0).Children(0).Name)
''	End Sub
'	
'	<Test> _
'	Public Sub TestGetDateRange
'		' プロパティファイルに設定された開始日と終了日の間の日付オブジェクトを、
'		' ひと月間隔で取得
'		Dim p As MyProperties = CreateProperties("20161001", "20161231")
'		
'		Dim r1 As List(Of DateTime) = UserRecordReader.GetDateRange(p)
'		Assert.AreEqual(3, r1.Count)
'		For i = 0 To r1.Count - 1
'			Assert.AreEqual(2016, r1(i).Year)
'			Assert.AreEqual(i + 10, r1(i).Month)
'		Next
'	End Sub
'	
'	Private Function CreateProperties(fromDate As String, toDate As String) As MyProperties
'		Dim p As New Properties(filepath)
'		p.Add(MyProperties.KEY_SELECTABLE_MIN_DATE, fromDate)
'		p.Add(MyProperties.KEY_SELECTABLE_MAX_DATE, toDate)
'		
'		Return New MyProperties(filepath)		
'	End Function
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
'End Class
