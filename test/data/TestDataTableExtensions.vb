'
' 日付: 2016/12/20
'
Imports NUnit.Framework

Imports System.Data
Imports System.Linq

<TestFixture> _
Public Class TestDataTableExtensions
  
  <Test> _
  Public Sub TestSkip
    Dim table As DataTable = CreateDefaultTable()
    
    AddData("john",   30, "England", table)
    AddData("george", 38, "America", table)
    AddData("gustav", 26, "Germen",  table)
    AddData("taro",   32, "Japan",   table)
    
    '
    ' 元のテーブルの３行目以降を抜き出しているかテスト
    '
    Dim subTable As DataTable = table.Skip(2)
    AssertEqualDataRow("gustav", 26, "Germen", subTable.Rows(0), "first row in subTable")
    AssertEqualDataRow("taro",   32, "Japan",  subTable.Rows(1), "second row in subTable")
  End Sub
  
  ''' <summary>
  ''' 行の値を判定する。
  ''' </summary>
  Private Sub AssertEqualDataRow(name As String, age As Integer, country As String, dataRow As DataRow, msg As String)
    Assert.AreEqual(name,    dataRow("name"),    msg)
    Assert.AreEqual(age,     dataRow("age"),     msg)
    Assert.AreEqual(country, dataRow("country"), msg)
  End Sub
  
  ''' <summary>
  ''' テスト用のテーブルを作成。
  ''' </summary>
  Private Function CreateDefaultTable() As DataTable
    Dim table As New DataTable
    table.Columns.Add(CreateColumn("name",    GetType(String)))
    table.Columns.Add(CreateColumn("age",     GetType(Integer)))
    table.Columns.Add(CreateColumn("country", GetType(String)))
    
    Return table
  End Function
  
  ''' <summary>
  ''' 列定義オブジェクトを簡単に作成するためのメソッド。
  ''' </summary>
  Private Function CreateColumn(name As String, type As Type) As DataColumn
    Dim col As New DataColumn
    col.ColumnName = name
    col.DataType   = type
    
    Return col
  End Function
  
  ''' <summary>
  ''' テーブルに行データを追加する。
  ''' </summary>
  Private Sub AddData(name As String, age As Integer, country As String, table As DataTable)
    Dim dataRow As DataRow = table.NewRow
    dataRow("name")    = name
    dataRow("age")     = age
    dataRow("country") = country
    
    table.Rows.Add(dataRow)
  End Sub
End Class
