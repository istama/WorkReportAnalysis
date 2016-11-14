''
'' 日付: 2016/06/22
''
'Imports NUnit.Framework
'Imports Moq
'Imports NSubstitute
'Imports Common.COM
'
'Imports System.Data
'
'' NSubstituteの使い方
'' http://nsubstitute.github.io/help/set-return-value/
'		
'<TestFixture> _
'Public Class TestExcelReader
'	Private mock As IExcel = Substitute.For(Of IExcel)()	
'	Private r As New ExcelReader(mock)
'	
'	' 木構造の列データを作成
'	Private Function CreateCols() As ColumnNode
'		Dim n As New ColumnNode("col1", "A", GetType(String))
'		n.AddChild("col2", "B", GetType(String)).AddChild("col3", "C", GetType(String))
'		n.AddChild("col4", "D", GetType(String))
'		
'		Return n
'	End Function
'	
'	<Test> _
'	Public Sub TestRead
'		' Excelを読み込みDataRowに格納する
'		Dim n As ColumnNode = CreateCols()
'		Dim table As DataTable = ColumnNode.CreateDataTable(n)
'		Dim d As DataRow = table.NewRow
'		
'		Dim filePath = "file.xls"
'		Dim sheetName = "sheet1"
'
'		' Excelのモックを作成
'		' Read()に指定した引数が渡された時に、Returns()で指定した値を返す
'		' モックメソッドの引数にオブジェクトを渡す場合は、必ずEquals()をオーバーライドすること
'		mock.Read(filePath, sheetName, Cell.Create(1, "A")).Returns("a")
'		mock.Read(filePath, sheetName, Cell.Create(1, "B")).Returns("")
'		mock.Read(filePath, sheetName, Cell.Create(1, "C")).Returns("c")
'		mock.Read(filePath, sheetName, Cell.Create(1, "D")).Returns("d")
'		
'		r.Read(filePath, sheetName, 1, n, d)
'		Assert.AreEqual("a", d("col1"))
'		Assert.AreEqual("", d("col2"))
'		Assert.AreEqual("", d("col3"))  ' 空文字列を返す要素の子要素は読み込まれない
'		Assert.AreEqual("d", d("col4"))
'		
'		' 同メソッドのオーバーライドを実行
'		Dim d2 As DataRow = table.NewRow
'		r.Read(filePath, sheetName, 1, n.Children, d2)
'		Assert.AreEqual("", d2("col2"))
'		Assert.AreEqual("", d2("col3"))  ' 空文字列を返す要素の子要素は読み込まれない
'		Assert.AreEqual("d", d2("col4"))		
'		
'		
''		m.
''		Setup(Sub(x) x.Write(It.IsAny(Of ExcelData))).
''		Callback(Of ExcelData)(
''			Sub(e)
''			Assert.AreEqual("1", e.WrittenText)
''			Assert.AreEqual("test.xls", e.filepath)
''			Assert.AreEqual("sheet1", e.sheetName)
''			Assert.AreEqual(1, e.cell.Row)
''			Assert.AreEqual("A", e.Cell.Col)
''			End Sub
''		)
''		
''		Dim w As New ExcelWriter(m.Object)
''		w.Init
''		w.AsyncWrite("1", "test.xls", "sheet1", Cell.Create(1,1))
''		m.Verify(Sub(x) x.Write(It.IsAny(Of ExcelData)()), Times.Once)
''		
''		w.Quit
'		
'' NSubsutituteによるテスト。
'' Subsutituteから生成されるモックはオリジナルの実装がそのまま実行されてしまう。
'' interfaceの実装でないクラスの場合、オリジナルの実装が足かせになる場合も。
''		Dim mock As IExcel = Substitute.For(Of IExcel)()
''		Dim w As New ExcelWriter(mock)
''		w.Init
''		w.AsyncWrite("1", "test.xls", "sheet1", Cell.Create(1,1))
''		mock.Received().Write(New ExcelData("1", "test.xls", "sheet1", Cell.Create(1,1)))
'		
''		w.AsyncWrite("2", "test.xls", "sheet1", Cell.Create(1,1))
''		w.AsyncWrite("3", "test.xls", "sheet1", Cell.Create(2,2))
''		w.AsyncWrite("4", "test.xls", "sheet1", Cell.Create(2,2))
''		w.AsyncWrite("5", "test.xls", "sheet1", Cell.Create(2,2))
''		w.AsyncWrite("6", "test.xls", "sheet2", Cell.Create(2,2))
'
'		
''		mock.ReceivedWithAnyArgs().Write("2", "test.xls", "sheet1", Cell.Create(1,1))
''		mock.ReceivedWithAnyArgs().Write(New ExcelData("3", "test.xls", "sheet1", Cell.Create(2,2)))
''		mock.Received().Write(New ExcelData("1", "test.xls", "sheet1", Cell.Create(1,1)))
''		w.AsyncWrite("2", "test.xls", "sheet1", Cell.Create(1,1))
''		mock.Received().Write(New ExcelData("2", "test.xls", "sheet1", Cell.Create(1,1)))
''		
''		w.AsyncWrite("3", "test.xls", "sheet1", Cell.Create(2,2))
''		mock.Received().Write(New ExcelData("3", "test.xls", "sheet1", Cell.Create(2,2)))
''		w.AsyncWrite("4", "test.xls", "sheet1", Cell.Create(2,2))
''		w.AsyncWrite("5", "test.xls", "sheet1", Cell.Create(2,2))
''		mock.Received().Write(New ExcelData("5", "test.xls", "sheet1", Cell.Create(2,2)))
''		w.AsyncWrite("6", "test.xls", "sheet2", Cell.Create(2,2))
''		mock.Received().Write(New ExcelData("6", "test.xls", "sheet1", Cell.Create(2,2)))
''		w.Quit
'	End Sub
'End Class
