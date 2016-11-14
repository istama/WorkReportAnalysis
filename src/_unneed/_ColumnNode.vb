''
'' 日付: 2016/06/22
''
'Imports System.Data
'Imports Common.COM
'
'''' <summary>
'''' 読み込むExcelの列とその列名を持つクラス。
'''' 複数のこのクラスのインスタンスを木構造に結合することができる。
'''' 列に対して処理を行うときに、子ノードの処理を親ノードの状態に依存させることなどができる。
'''' </summary>
'Public Class ColumnNode
'	''' 列名	
'	Private _name As String
'	Public ReadOnly Property Name As String
'		Get 
'			Return _name
'		End Get
'	End Property
'	
'	''' Excel上の列
'	Private _col As String
'	Public ReadOnly Property Col As String
'		Get
'			Return _col
'		End Get
'	End Property
'	
'	Private _valueType As Type
'	Public ReadOnly Property ValueType As Type
'		Get
'			Return _valueType
'		End Get
'	End Property
'	
'	''' 子ノード
'	Private _children As List(Of ColumnNode)
'	Public ReadOnly Property Children As List(Of ColumnNode)
'		Get
'			Return _children
'		End Get
'	End Property
'	
'	Public Sub New(name As String, col As String, valueType As Type)
'		If name Is Nothing      Then Throw New ArgumentNullException("name is null")
'		If col Is Nothing       Then Throw New ArgumentNullException("col is null")
'		If valueType Is Nothing Then Throw New ArgumentNullException("valueType is null")
'		
'		If Not Cell.ValidColumn(col) Then
'			Throw New ArgumentException("col is a invalid value")
'		End If
'		
'		If valueType <> GetType(String) AndAlso valueType <> GetType(Double) AndAlso valueType <> GetType(Integer) Then
'			Throw New ArgumentException("the type that be allowed to the A are ""String"", ""Double"" and ""Integer"".")
'		End If
'		
'		_name      = name
'		_col       = col
'		_valueType = valueType
'		_children  = Nothing
'	End Sub
'	
'	''' <summary>
'	''' 子ノードを生成し、追加する。
'	''' 戻り値は生成した子ノードを返す。
'	''' </summary>
'	Public Function AddChild(name As String, col As String, valueType As Type) As ColumnNode
'		Dim node As New ColumnNode(name, col, valueType)
'		
'		If _children Is Nothing Then
'			_children = New List(Of ColumnNode)
'		End If		
'		_children.Add(node)
'		
'		Return node
'	End Function
'	
'	Public OVerrides Function ToString() As String
'		Dim str As String = String.Empty
'		If _children IsNot Nothing
'			For Each c In _children
'				If String.IsNullOrEmpty(str) Then
'					str = c.ToString
'				Else
'					str = str & ", " & c.ToString
'				End If
'			Next
'			
'			Return _name & ", " & str
'		Else
'			Return _name
'		End If
'	End Function
'	
'	''' <summary>
'	''' 列のデータ構造からデータテーブルを作成する。
'	''' </summary>
'	Public Shared Function CreateDataTable(colNode As ColumnNode) As DataTable
'		If colNode Is Nothing Then
'			Throw New ArgumentNullException("colNode is null")
'		End If
'		
'		Dim table As New DataTable()
'		table.Columns.Add(colNode.Name)
'		If colNode.Children IsNot Nothing Then
'			CreateColumns(table, colNode.Children)
'		End If
'		
'		Return table		
'	End Function
'	
'	''' <summary>
'	''' 列のデータ構造からデータテーブルを作成する。
'	''' </summary>
'	Public Shared Function CreateDataTable(colList As List(Of ColumnNode)) As DataTable
'		If colList Is Nothing Then
'			Throw New ArgumentNullException("colList is null")
'		End If
'		
'		Dim table As New DataTable()
'		CreateColumns(table, colList)
'		
'		Return table			
'	End Function
'	
'	Public Shared Function CreateColumn(node As ColumnNode) As DataColumn
'		Dim col As New DataColumn
'		col.DataType = node.ValueType
'		col.ColumnName = node.Name
'		col.AutoIncrement = False
'		
'		Return col
'	End Function
'	
'	Private Shared Sub CreateColumns(table As DataTable, colList As List(Of ColumnNode))
'		For Each node In colList
'			Dim col As DataColumn = CreateColumn(node)
'			table.Columns.Add(col)
'			If node.Children IsNot Nothing
'				CreateColumns(table, node.Children)
'			End If
'		Next
'	End Sub
'End Class
