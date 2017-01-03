'
' 日付: 2016/11/23
'
Imports System.Data
Imports System.Linq
Imports Common.Extensions

Public Module DataTableExtensions
  
  ''' <summary>
  ''' データテーブルの値をコピーする。
  ''' ただし、コピーされるのは２つのテーブルが共通して持っている名前と型の列の値のみ。
  ''' </summary>
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Sub CopyTo(_from As DataTable, _to As DataTable)
    ' ２つのテーブルのうち少ないほうの行数を取得
    Dim minRowCount As Integer = Math.Min(_from.Rows.Count, _to.Rows.Count)
    
    If minRowCount = 0 Then
      Return
    End If
    
    ' ２つのテーブルが共通して持っている列を取得
    Dim commonCols = _from.Rows(0).CommonColumns(_to.Rows(0))
    
    ' コピーする
    For idx As Integer = 0 To minRowCount - 1
      Dim fromRow As DataRow = _from.Rows(idx)
      Dim toRow   As DataRow = _to.Rows(idx)
      For Each col As DataColumn In commonCols
        toRow(col.ColumnName) = fromRow(col.ColumnName)
      Next
    Next
  End Sub
  
  ''' <summary>
  ''' データテーブルの行をコピーし、追加する。
  ''' ただし、コピーされるのは２つのテーブルが共通して持っている名前と型の列の値のみ。
  ''' </summary>
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Sub WriteTo(_from As DataTable, _to As DataTable)
    If _from.Rows.Count() = 0 Then Return
    
    ' ２つのテーブルが共通して持っている列を取得
    Dim r As DataRow
    If _to.Rows.Count > 0 Then
      r = _to.Rows(0)
    Else
      r = _to.NewRow
    End If
    Dim commonCols = _from.Rows(0).CommonColumns(r)
    
    ' コピーする
    For idx As Integer = 0 To _from.Rows.Count - 1
      Dim fromRow As DataRow = _from.Rows(idx)
      Dim toRow   As DataRow = _to.NewRow
      For Each col As DataColumn In commonCols
        toRow(col.ColumnName) = fromRow(col.ColumnName)
      Next
      _to.Rows.Add(toRow)
    Next
  End Sub
  
  ''' <summary>
  ''' 指定した行数までのデータを新たなテーブルにセットして返す。
  ''' </summary>
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Function Take(_from As DataTable, rowcnt As Integer) As DataTable
    Dim newTable As DataTable = _from.Clone
    
    For Each row As DataRow In _from.AsEnumerable().Take(rowcnt)
      newTable.ImportRow(row)
    Next
    
    Return newTable
  End Function
  
  ''' <summary>
  ''' 指定した行数以降のデータのみを新たなテーブルにセットして返す。
  ''' </summary>
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Function Skip(_from As DataTable, rowcnt As Integer) As DataTable
    Dim newTable As DataTable = _from.Clone
    
    For Each row As DataRow In _from.AsEnumerable().Skip(rowcnt)
      newTable.ImportRow(row)
    Next
    
    Return newTable
  End Function
  
  ''' <summary>
  ''' 指定した列の合計値をDouble型で求める。
  ''' </summary>
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Function SumByDouble(table As DataTable, colName As String) As Double
    Return _
      table.AsEnumerable().
        Where(Function(row) row.IsNull(colName) = False).
        Select(Function(row) row.Field(Of Double)(colName)).Sum()
  End Function
  
  ''' <summary>
  ''' 指定した列の合計値をDouble型で求める。
  ''' ただし条件に当てはまらなかった列の値合計に含めない。
  ''' </summary>
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Function SumByDouble(table As DataTable, colName As String, predicate As Func(Of DataRow, Boolean)) As Double
    Return _
      table.AsEnumerable().
        Where(Function(row) row.IsNull(colName) = False).
        Where(Function(row) predicate(row)).
        Select(Function(row) row.Field(Of Double)(colName)).Sum()
  End Function
  
  ''' <summary>
  ''' 指定した列の合計値をInt型で求める。
  ''' </summary>
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Function SumByInteger(table As DataTable, colName As String) As Integer
    Return _
      table.AsEnumerable().
        Where(Function(row) row.IsNull(colName) = False).
        Select(Function(row) row.Field(Of Integer)(colName)).Sum()
  End Function
  
  ''' <summary>
  ''' 指定した列の合計値をInt型で求める。
  ''' ただし条件に当てはまらなかった列の値合計に含めない。
  ''' </summary>
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Function SumByInteger(table As DataTable, colName As String, predicate As Func(Of DataRow, Boolean)) As Integer
    Return _
      table.AsEnumerable().
        Where(Function(row) row.IsNull(colName) = False).
        Where(Function(row) predicate(row)).
        Select(Function(row) row.Field(Of Integer)(colName)).Sum()
  End Function
  
  ''' <summary>
  ''' 各列の合計値をおさめたDataRowを返す。
  ''' ただし、IntegerとDouble型以外の列の合計は求めない。
  ''' </summary>
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Function SumRow(table As DataTable) As DataRow
    Dim sum As DataRow = table.NewRow
    
    For Each col As DataColumn In table.Columns
      If col.DataType = GetType(Integer) Then
        sum(col.ColumnName) = table.SumByInteger(col.ColumnName)
      ElseIf col.DataType = GetType(Double)
        sum(col.ColumnName) = table.SumByDouble(col.ColumnName)
      End If
    Next
    
    Return sum
  End Function
  
  ''' <summary>
  ''' 各列の合計値を引数のDataRowにセットする。
  ''' ただし、渡したテーブルと行が共通して持っている列以外の列はスルーする。
  ''' また、IntegerとDouble型以外の列の合計は求めない。
  ''' </summary>
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Sub SumRow(table As DataTable, resultRow As DataRow)
    Dim row As DataRow = table.NewRow
    
    For Each col As DataColumn In row.CommonColumns(resultRow)
      If col.DataType = GetType(Integer) Then
        resultRow(col.ColumnName) = table.SumByInteger(col.ColumnName)
      ElseIf col.DataType = GetType(Double)
        resultRow(col.ColumnName) = table.SumByDouble(col.ColumnName)
      End If
    Next
  End Sub
  
  ''' <summary>
  ''' CSVに変換する。
  ''' </summary>
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Iterator Function ToCSV(table As DataTable) As IEnumerable(Of String)
    Dim colNames = table.Columns.ToEnumerable().Select(Function(col) col.ColumnName)
    
    ' 列名のCSVを作成
    Yield String.Join(","c, colNames)
    
    ' 各行の値のCSVを作成
    For Each row As DataRow In table.Rows
      Dim values = 
        colNames.Select(
          Function(name)
            If row.IsNull(name) Then
              Return String.Empty
            Else
              Return row(name).ToString
            End If
          End Function)
      Yield String.Join(","c, values)
    Next
  End Function
End Module
