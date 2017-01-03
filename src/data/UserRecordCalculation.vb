'
' 日付: 2016/12/20
'
Imports System.Data
Imports System.Linq

Public Class UserRecordCalculation
  
  ''' <summary>
  ''' 各列の合計値を求め、新しい行にセットして返す。
  ''' ただし、作業時間の値が０の行では同じ行の作業件数の値を合計値に含めない。
  ''' </summary>
  Public Shared Function SumRowExceptedUnfilled(table As DataTable, columnsInfo As UserRecordColumnsInfo) As DataRow
    Dim sumRow As DataRow = table.NewRow
    
    For Each cols As WorkItemColumnsInfo In columnsInfo.WorkItems
      Dim cntColName  As String = cols.WorkCountColInfo.name
      Dim timeColName As String = cols.WorkTimeColInfo.name
      
      ' 作業件数の合計値を求める
      If String.IsNullOrWhiteSpace(cntColName) = False Then
        Dim cntSum As Integer = _
          table.SumByInteger(
            cntColName,
            Function(row) 
              ' 作業時間の列が存在しない、もしくは作業時間の列の値が０より大きい場合、作業件数の値を合計値に含める
              Return _
                String.IsNullOrWhiteSpace(timeColName) OrElse _
                (row.IsNull(timeColName) = False AndAlso row.Field(Of Double)(timeColName) > 0)
            End Function)
        
        sumRow(cntColName)  = cntSum
      End If
      
      ' 作業時間の合計値を求める
      If String.IsNullOrWhiteSpace(timeColName) = False Then
        Dim timeSum As Double = table.SumByDouble(timeColName)
        sumRow(timeColName) = timeSum
      End If
    Next
    
    Return sumRow
  End Function

End Class
