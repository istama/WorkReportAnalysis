''
'' 日付: 2016/07/07
''
'Imports NUnit.Framework
'
'Imports System.Data
'Imports Common.IO
'Imports Common.Util
'
'<TestFixture> _
'Public Class TestUserRecord
'	Private format As String = "month{0}"	
'	Private filepath As String = "./test/test.properties"
'	
'	Private Function CreateUserRecord As UserRecord
'		Dim p As New MyProperties(filepath)
'		Dim record As New UserRecord(p)
'		Dim cols As DataColumnCollection = record.TableColumn
'		
'		For m = 10 To 12
'			Dim table As DataTable = record.CreateTable(m)
'			For d = 1 To DateTime.DaysInMonth(2016, m)
'				Dim r As DataRow = table.NewRow
'				r(0)  = d * 10
'				r(1)  = d
'				r(2)  = 10
'				r(3)  = d * 7
'				r(4)  = d
'				r(5)  = 7
'				table.Rows.Add(r)
'			Next
'		Next
'		
'		Return record
'	End Function
'	
'	<Test> _
'	Public Sub TestGetTable
'		' 指定した名前のテーブルを取得
'		Dim record As UserRecord = CreateUserRecord
'		Dim cols As DataColumnCollection = record.TableColumn
'		
'		Dim table1 As DataTable = record.GetTable(10)
'		Assert.AreEqual(31, table1.Rows.Count)
'		Assert.AreEqual("10", table1.Rows(0)(0))
'		
'		Dim table2 As DataTable = record.GetTable(11)
'		Assert.AreEqual(30, table2.Rows.Count)
'		Assert.AreEqual("14", table2.Rows(1)(3))
'		
'		Dim table3 As DataTable = record.GetTable(12)
'		Assert.AreEqual(31, table3.Rows.Count)
'		Assert.AreEqual("30", table3.Rows(2)(0))
'	End Sub
'	
'	<Test> _
'	Public Sub TestGetTable2
'		' 指定した期間のデータを抜き出す
'		Dim record As UserRecord = CreateUserRecord
'		Dim cols As DataColumnCollection = record.TableColumn
'		
'		Dim table1 As DataTable = record.GetTable(#10/05/2016#, #10/15/2016#)
'		Assert.AreEqual(11, table1.Rows.Count)
'		Assert.AreEqual("50", table1.Rows(0)(cols(0).ToString))
'		Assert.AreEqual("105", table1.Rows(10)(cols(3).ToString))
'		
'		Dim table2 As DataTable = record.GetTable(#10/20/2016#, #11/10/2016#)
'		Assert.AreEqual(22, table2.Rows.Count)
'		Assert.AreEqual("200", table2.Rows(0)(cols(0).ToString))
'		Assert.AreEqual("7", table2.Rows(12)(cols(3).ToString))
'		Assert.AreEqual("100", table2.Rows(21)(cols(0).ToString))
'		
'		Dim table3 As DataTable = record.GetTable(#10/01/2016#, #12/31/2016#)
'		Assert.AreEqual(92, table3.Rows.Count)
'		Assert.AreEqual("10", table3.Rows(0)(cols(0).ToString))
'		Assert.AreEqual("217", table3.Rows(30)(cols(3).ToString))
'		Assert.AreEqual("10", table3.Rows(31)(cols(0).ToString))
'		Assert.AreEqual("210", table3.Rows(60)(cols(3).ToString))
'		Assert.AreEqual("10", table3.Rows(61)(cols(0).ToString))
'		Assert.AreEqual("217", table3.Rows(91)(cols(3).ToString))
'	End Sub
'	
'	<Test> _
'	Public Sub TestCreateColumnNode
'		' プロパティファイルに設定された列と列名から、Excelにアクセスするための列ノードを作成する
'		Dim pp As New Properties(filepath)
'		
'		' これらの項目はプロパティ値がセットされてないため、列ノードには含まれない
'		pp.Add(MyProperties.KEY_ITEM_NAME & "2", "")
'		pp.Add(MyProperties.KEY_WORKCOUNT_COL_OF_ITEM & "4", "")
'		pp.Add(MyProperties.KEY_WORKTIME_COL_OF_ITEM & "6", "")
'		
'		Dim p As New MyProperties(filepath)
'		Dim n As ColumnNode = UserRecord.CreateColumnNode(p)
'		
'		Assert.AreEqual("出勤日", n.Name)
'		Dim idx As Integer = 0
'		For i = 1 To p.GetItemValuesList.Count
'			' 列ノードに含まれるインデックス
'			If i = 1 OrElse i = 3 OrElse i = 5 OrElse i = 7 Then
'				CheckWorkItemColumn(p.GetItemName(i), n.Children(idx))
'				idx += 1
'			End If
'		Next
'		Assert.AreEqual(p.GetValue(MyProperties.KEY_NOTE_NAME), n.Children(n.Children.Count - 1).Name)
'	End Sub
'	
'	Private Sub CheckWorkItemColumn(colName As String, node As ColumnNode)
'		Assert.AreEqual(colName & vbCrLf & "件数",     node.Name)
'		Assert.AreEqual(colName & vbCrLf & "作業時間", node.Children(0).Name)
'		Assert.AreEqual(colName & vbCrLf & "生産性",   node.Children(0).Children(0).Name)
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
'End Class
