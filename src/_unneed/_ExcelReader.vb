''
'' 日付: 2016/06/20
''
'Imports System.Data
'Imports Common.COM
'Imports System.IO
'
'''' <summary>
'''' Excelを読み込むクラス。
'''' このクラスではファイルを読み込む際、読み込む列を木構造状のデータに格納してアクセスする。
'''' 親要素となる列にアクセスし、仮にそれが空文字を返した場合、
'''' 以降の子要素の列はファイルにアクセスせず、空文字を返すようにする。
'''' そうすることで、ある列のデータが別の列のデータに依存する場合、余計なアクセスを減らし、
'''' パフォーマンスが向上する。
'''' </summary>
'Public Class ExcelReader
'	Public Structure GridColumn
'		Dim Name As String
'		Dim Col As String
'	End Structure
'	
'	Public Structure ColumnTree
'		Dim GridColumn As GridColumn
'		Dim Children As List(Of ColumnTree)
'	End Structure
'	
'	' Excel	
'	Private excel As IExcel
'	
'	Public Sub New(excel As IExcel)
'		Me.excel = excel
'	End Sub
'	
'	Public Sub Init()
'		excel.init
'	End Sub
'	
'	Public Sub Quit()
'		excel.Quit
'	End Sub
'	
'	
'	''' <summary>
'	''' Excelファイルを1行分読み込み、dataRowに格納する。
'	''' </summary>
'	''' <param name="row">アクセスする行</param>
'	'''　<param name="cols">アクセスする列を木構造状に格納したオブジェクト</param>
'	''' <param name="dataRow">アクセスしたデータを格納するオブジェクト</param> 
'	Public Sub Read(filePath As String, sheetName As String, row As Integer, cols As ColumnNode, dataRow As DataRow)
'		Read(filePath, sheetName, row, cols, True, dataRow)
'	End Sub
'	
'	''' <summary>
'	''' Excelファイルを1行分読み込み、dataRowに格納する。
'	''' </summary>
'	''' <param name="row">アクセスする行</param>
'	'''　<param name="cols">アクセスする列を木構造状に格納したオブジェクト</param>
'	''' <param name="dataRow">アクセスしたデータを格納するオブジェクト</param> 
'	Public Sub Read(filePath As String, sheetName As String, row As Integer, colList As List(Of ColumnNode), dataRow As DataRow)
'		Read(filePath, sheetName, row, colList, True, dataRow)
'	End Sub
'	
'	Private Sub Read(filePath As String, sheetName As String, row As Integer, colNode As ColumnNode, readChild As Boolean, dataRow As DataRow)
'		Dim text As String = String.Empty
'		
'		' 親ノードが子ノードはファイルを読み込むと判断した場合のみ、ファイルアクセスする
'		If readChild Then
'			'text = excel.Read(filepath, sheetName, Cell.Create(row, colNode.Col))
'			'text = ((row + row) * 17 * (Convert.ToInt32(colNode.Col.ToCharArray()(0)) - Convert.ToInt32("A"c) + 1) Mod 30).ToString
'			text = ((row + (Convert.ToInt32(colNode.Col.ToCharArray()(0)) - Convert.ToInt32("A"c) + 1)) Mod 18).ToString
'			'Console.WriteLine("read " & filepath & " " & sheetName & " " & colNode.Name & " " & colNode.Col & " " & row & " " & text)
'			' 読み込んだ文字列からから文字だった場合、子要素は読み込まない
'			If String.IsNullOrEmpty(text) Then
'				readChild = False
'			End If
'		End If
'		
'		Select Case colNode.ValueType
'			Case GetType(Double)
'				Dim v As Double
'				If Double.TryParse(text, v) AndAlso v <> 0 Then
'					dataRow(colNode.Name) = v
'				End If				
'				
'			Case GetType(Integer)
'				Dim v As Integer
'				If Integer.TryParse(text, v) AndAlso v <> 0 Then
'					dataRow(colNode.Name) = v
'				End If				
'			Case Else
'				dataRow(colNode.Name) = text
'		End Select
'				
'		If colNode.Children IsNot Nothing AndAlso colNode.Children.Count > 0 Then
'			Read(filePath, sheetName, row, colNode.Children, readChild, dataRow)
'		End If		
'	End Sub
'	
'	Private Sub Read(filePath As String, sheetName As String, row As Integer, colList As List(Of ColumnNode), readChild As Boolean, dataRow As DataRow)
'		For Each col In colList
'			Read(filePath, sheetName, row, col, readChild, dataRow)
'		Next
'	End Sub
'	
'End Class
