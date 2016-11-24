'
' 日付: 2016/11/23
'
Imports System.Data
Imports System.Linq
Imports Common.Extensions

Public Module DataTableExtensions
  
  ''' <summary>
  ''' 引数のテーブルの各列の値の合計を、引数の行オブジェクトにセットする。
  ''' </summary>
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Sub Sum(table As DataTable, sumRow As DataRow, columnsInfo As UserRecordColumnsInfo)
    columnsInfo.WorkItemList.ForEach(
      Sub(item)
        calcTotal(table, sumRow, item.WorkCountColInfo)
        calcTotal(table, sumRow, item.WorkTimeColInfo)
      End Sub)
  End Sub
  
  ''' <summary>
  ''' 引数のテーブルの各列の値の合計を、引数の行オブジェクトにセットする。
  ''' ただし、作業時間が０である作業件数の値は合計に含めない。 
  ''' </summary>
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Sub SumExceptWorkCountOfZeroWorkTimeIs(table As DataTable, sumRow As DataRow, columnsInfo As UserRecordColumnsInfo)
    columnsInfo.WorkItemList.ForEach(
      Sub(item)
        Dim cntColName As String = item.WorkCountColInfo.Name
        Dim timeColName As String = item.WorkTimeColInfo.Name
        Dim existsCntCol As Boolean = Not String.IsNullOrEmpty(cntColName)
        Dim existsTimeCol As Boolean = Not String.IsNullOrEmpty(timeColName)
        
        If existsCntCol AndAlso (Not table.Columns.Contains(cntColName) OrElse Not sumRow.HasColumn(cntColName)) Then
          Throw New ArgumentException("テーブルに存在しない列名です。 / " & cntColName)          
        End If
        
        If existsTimeCol AndAlso (Not table.Columns.Contains(timeColName) OrElse Not sumRow.HasColumn(timeColName)) Then
          Throw New ArgumentException("テーブルに存在しない列名です。 / " & timeColName)          
        End If
        
        ' 作業件数の合計を求める
        ' 作業時間が０の場合は合計に含めない
        If existsCntCol Then
          Dim sum As Integer =
            table.Rows.ToEnumerable().
              Where(Function(row) Not row.IsNull(cntColName)).
              Where(Function(row) Not existsTimeCol OrElse Not row.IsNull(timeColName) AndAlso row.ToDouble(timeColName) > 0.0).
              Sum(Function(row) row.ToInt(cntColName))
          
          sumRow(cntColName) = sum
        End If
        
        ' 作業時間の合計を求める
        If existsTimeCol Then
          Dim sum As Double =
            table.Rows.ToEnumerable().
              Where(Function(row) Not row.IsNull(timeColName)).
              Sum(Function(row) row.ToDouble(timeColName))
          
          sumRow(timeColName) = sum
        End If
      End Sub)    
  End Sub
  
  Private Sub calcTotal(table As DataTable, sumRow As DataRow, colInfo As ColumnInfo)
    If String.IsNullOrEmpty(colInfo.name) Then Return    
    
    If Not table.Columns.Contains(colInfo.Name) OrElse Not sumRow.HasColumn(colInfo.Name) Then
      Throw New ArgumentException("テーブルに存在しない列名です。 / " & colInfo.Name)
    End If
    
    If colInfo.type = GetType(Integer) Then
      Dim sum As Integer = 
        table.Rows.ToEnumerable().
          Where(Function(row) Not row.IsNull(colInfo.name)).
          Sum(Function(row) row.ToInt(colInfo.name))
      
      sumRow(colInfo.name) = sum
    ElseIf colInfo.type = GetType(Double) Then
      Dim sum As Double = 
        table.Rows.ToEnumerable().
          Where(Function(row) Not row.IsNull(colInfo.name)).
          Sum(Function(row) row.ToDouble(colInfo.name))
      
      sumRow(colInfo.name) = sum      
    End If
  End Sub
  
End Module
