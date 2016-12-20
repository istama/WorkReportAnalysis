'
' 日付: 2016/12/20
'
Imports NUnit.Framework

Imports System.Data
Imports System.Linq

<TestFixture> _
Public Class TestDataTableExtensions
  
  <Test> _
  Public Sub TestTake
    Dim table As DataTable = CreateDefaultTable()
    
    AddData("john",   30, "England", 58.4, table)
    AddData("george", 38, "America", 71.6, table)
    AddData("gustav", 26, "Germen",  60.7, table)
    AddData("taro",   32, "Japan",   68.3, table)
    
    '
    ' 元のテーブルの２行目までを抜き出しているかテスト
    '
    Dim subTable As DataTable = table.Take(2)
    Assert.AreEqual(2, subTable.Rows.Count, "row size")
    AssertEqualDataRow("john",   30, "England", 58.4, subTable.Rows(0), "first row in subTable")
    AssertEqualDataRow("george", 38, "America", 71.6, subTable.Rows(1), "second row in subTable")    
  End Sub
  
  <Test> _
  Public Sub TestSkip
    Dim table As DataTable = CreateDefaultTable()
    
    AddData("john",   30, "England", 58.4, table)
    AddData("george", 38, "America", 71.6, table)
    AddData("gustav", 26, "Germen",  60.7, table)
    AddData("taro",   32, "Japan",   68.3, table)
    
    '
    ' 元のテーブルの３行目以降を抜き出しているかテスト
    '
    Dim subTable As DataTable = table.Skip(2)
    AssertEqualDataRow("gustav", 26, "Germen", 60.7, subTable.Rows(0), "first row in subTable")
    AssertEqualDataRow("taro",   32, "Japan",  68.3, subTable.Rows(1), "second row in subTable")
  End Sub
  
  <Test> _
  Public Sub TestSumByDouble
    Dim table As DataTable = CreateDefaultTable()
    
    AddData("john",   30, "England", 58.4, table)
    AddData("george", 38, "America", 71.6, table)
    AddData("gustav", 26, "Germen",  60.7, table)
    AddData("taro",   32, "Japan",   68.3, table)
    
    '
    ' 指定した列の合計を求める
    '
    Assert.AreEqual(259.0, table.SumByDouble("weight"))
    
    '
    ' ageが30以上のみの合計を求める
    '
    Assert.AreEqual(
      139.9,
      Math.Round(table.SumByDouble("weight", Function(row) row.Field(Of Integer)("age") > 30)), 1)
  End Sub
  
  <Test> _
  Public Sub TestSumByInteger
    Dim table As DataTable = CreateDefaultTable()
    
    AddData("john",   30, "England", 58.4, table)
    AddData("george", 38, "America", 71.6, table)
    AddData("gustav", 26, "Germen",  60.7, table)
    AddData("taro",   32, "Japan",   68.3, table)
    
    '
    ' 指定した列の合計を求める
    '
    Assert.AreEqual(126, table.SumByInteger("age"))
    
    '
    ' weightが65.0以上のみの合計を求める
    '
    Assert.AreEqual(
      70, table.SumByInteger("age", Function(row) row.Field(Of Double)("weight") > 65.0), 1)
  End Sub
  
  <Test> _
  Public Sub TestSumRow
    Dim table As DataTable = CreateDefaultTable()
    
    AddData("john",   30, "England", 58.4, table)
    AddData("george", 38, "America", 71.6, table)
    AddData("gustav", 26, "Germen",  60.7, table)
    AddData("taro",   32, "Japan",   68.3, table)
    
    '
    ' 各列の合計値を収めた列を返しているかテスト
    '
    Dim sumRow As DataRow = table.SumRow()
    Assert.AreEqual(126,   sumRow("age"),    "sum of age")
    Assert.AreEqual(259.0, sumRow("weight"), "sum of weight")
  End Sub
  
  ''' <summary>
  ''' 行の値を判定する。
  ''' </summary>
  Private Sub AssertEqualDataRow(name As String, age As Integer, country As String, weight As Double, dataRow As DataRow, msg As String)
    Assert.AreEqual(name,    dataRow("name"),    msg)
    Assert.AreEqual(age,     dataRow("age"),     msg)
    Assert.AreEqual(country, dataRow("country"), msg)
    Assert.AreEqual(weight,  dataRow("weight"),  msg)
  End Sub
  
  ''' <summary>
  ''' テスト用のテーブルを作成。
  ''' </summary>
  Private Function CreateDefaultTable() As DataTable
    Dim table As New DataTable
    table.Columns.Add(CreateColumn("name",    GetType(String)))
    table.Columns.Add(CreateColumn("age",     GetType(Integer)))
    table.Columns.Add(CreateColumn("country", GetType(String)))
    table.Columns.Add(CreateColumn("weight",  GetType(Double)))
    
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
  Private Sub AddData(name As String, age As Integer, country As String, weight As Double, table As DataTable)
    Dim dataRow As DataRow = table.NewRow
    dataRow("name")    = name
    dataRow("age")     = age
    dataRow("country") = country
    dataRow("weight")  = weight
    
    table.Rows.Add(dataRow)
  End Sub
End Class
