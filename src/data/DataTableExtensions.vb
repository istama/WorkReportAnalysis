'
' 日付: 2016/11/23
'
Imports System.Data
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
        
        If existsCntCol AndAlso (Not table.Columns.Contains(cntColName) OrElse sumRow.Table.Columns.Contains(cntColName)) Then
          Throw New ArgumentException("テーブルに存在しない列名です。 / " & cntColName)          
        End If
        
        If existsTimeCol AndAlso (Not table.Columns.Contains(timeColName) OrElse sumRow.Table.Columns.Contains(timeColName)) Then
          Throw New ArgumentException("テーブルに存在しない列名です。 / " & timeColName)          
        End If
        
        table.Rows.Convert(
          Function(row)
            If existsTimeCol Then
              plusDouble(row, sumRow, timeColName)
            End If
            Return row
          End Function
        )._Where(
          Function(row) existsCntCol AndAlso (Not existsTimeCol OrElse DirectCast(row(timeColName), Double) > 0.0)
        ).ForEach(
          Function(row) plusInt(row, sumRow, cntColName)
        )
          
        
'        For Each row As DataRow In table.Rows
'          If existsTimeCol Then
'            If Not System.Convert.IsDBNull(row(timeColName)) Then
'              Dim time As Double = DirectCast(row(timeColName), Double)
'              If time > 0.0 Then
'                plusDouble(row, sumRow, timeColName)
'                
'                If existsCntCol Then
'                  plusInt(row, sumRow, cntColName)
'                End If
'              End If
'            End If
'          ElseIf existsCntCol Then
'            plusInt(row, sumRow, cntColName)
'          End If
'        Next
          
      End Sub)    
  End Sub
  
  Private Sub calcTotal(table As DataTable, sumRow As DataRow, colInfo As ColumnInfo)
    If Not String.IsNullOrEmpty(colInfo.Name) Then
      If table.Columns.Contains(colInfo.Name) AndAlso sumRow.Table.Columns.Contains(colInfo.Name) Then
        If colInfo.type = GetType(Integer) Then
          table.Rows.ForEach(Sub(row) plusInt(row, sumRow, colInfo.Name))
        ElseIf colInfo.type = GetType(Double)
          table.Rows.ForEach(Sub(row) plusDouble(row, sumRow, colInfo.Name))
        End If
      Else
        Throw New ArgumentException("テーブルに存在しない列名です。 / " & colInfo.Name)          
      End If
    End If    
  End Sub
  
  Private Function plusInt(valueRow As DataRow, addedRow As DataRow, colName As String) As Boolean
    If Not System.Convert.IsDBNull(valueRow(colName)) Then
      If System.Convert.IsDBNull(addedRow(colName)) Then
        addedRow(colName) = DirectCast(valueRow(colName), Integer)
      Else
        addedRow(colName) = DirectCast(addedRow(colName), Integer) + DirectCast(valueRow(colName), Integer)
      End If
      
      Return True
    Else
      Return False
    End IF
  End Function
  
  Private Function plusDouble(valueRow As DataRow, addedRow As DataRow, colName As String) As Boolean
    If Not System.Convert.IsDBNull(valueRow(colName)) Then
      If System.Convert.IsDBNull(addedRow(colName)) Then
        addedRow(colName) = DirectCast(valueRow(colName), Double)
      Else
        addedRow(colName) = DirectCast(addedRow(colName), Double) + DirectCast(valueRow(colName), Double)
      End If
      
      Return True
    Else
      Return False
    End IF
  End Function
End Module
