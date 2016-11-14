''
'' 日付: 2016/06/22
''
'Imports NUnit.Framework
'
'Imports System.Data
'
'<TestFixture> _
'Public Class TestColumnNode
'	
'	<Test> _
'	Public Sub TestConstruction
'		' コンストラクション時の値が正しいか判定する
'		Dim node As New ColumnNode("count", "A", GetType(String))
'		
'		Assert.AreEqual("count", node.Name)
'		Assert.AreEqual("A", node.Col)
'		Assert.IsNull(node.Children)
'	End Sub
'	
'	<Test> _
'	Public Sub TestAddChild
'		' 子ノードを追加する
'		Dim node As New ColumnNode("count", "A", GetType(String))
'		Dim child As ColumnNode = node.AddChild("time", "B", GetType(String))
'		
'		Assert.AreEqual("time", child.Name)
'		Assert.AreEqual("B", child.Col)
'		Assert.IsNull(child.Children)
'		Assert.IsTrue(node.Children(0) Is child)
'	End Sub
'	
'	<Test> _
'	Public Sub TestCreateDataTable
'		' 列のデータ構造からデータテーブルを作成する
'		Dim node As New ColumnNode("col1", "A", GetType(String))
'		node.AddChild("col2", "B", GetType(String)).AddChild("col3", "C", GetType(String))
'		node.AddChild("col4", "D", GetType(String))
'		
'		' 列ノードの順番とデータテーブルの列の順番が一致しているか判定
'		Dim table As DataTable = ColumnNode.CreateDataTable(node)
'		Dim cols As DataColumnCollection = table.Columns
'		
'		Assert.AreEqual("col1", cols(0).ColumnName)
'		Assert.AreEqual("col2", cols(1).ColumnName)
'		Assert.AreEqual("col3", cols(2).ColumnName)
'		Assert.AreEqual("col4", cols(3).ColumnName)
'		
'		' オーバーロードされた同メソッドのテスト
'		Dim table2 As DataTable = ColumnNode.CreateDataTable(node.Children)
'		Dim cols2 As DataColumnCollection = table2.Columns
'		
'		Assert.AreEqual("col2", cols2(0).ColumnName)
'		Assert.AreEqual("col3", cols2(1).ColumnName)
'		Assert.AreEqual("col4", cols2(2).ColumnName)	
'	End Sub
'End Class
