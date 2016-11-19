'
' 日付: 2016/10/18
'
Imports Common.COM
Imports Common.Util
Imports Common.Account
Imports System.Data
Imports Common.Extensions
Imports Common.IO

''' <summary>
''' ユーザのExcelデータを読み込むクラス。
''' </summary>
Public Class UserRecordReader

  
  Private ReadOnly properties As ExcelProperties  
  Private ReadOnly excel As ExcelReaderByColumnTree
  
  Private _cancel As Boolean = False
  
  Public Sub New(properties As ExcelProperties, excel As Excel3)
    If properties Is Nothing Then Throw New ArgumentNullException("properties is null")
    If excel      Is Nothing Then Throw New ArgumentNullException("excel is null")
    
    Me.properties     = properties
    Me.excel          = New ExcelReaderByColumnTree(excel)
  End Sub
  
  Public Sub Cancel
    Me._cancel = True
  End Sub
  
  ''' <summary>
  ''' 指定したユーザのレコードを読み込む。
  ''' </summary>
  Public Sub Read(userRecord As UserRecord)
    If userRecord Is Nothing Then Throw New ArgumentNullException("userRecord is null")
    
    Dim filepath As String = String.Format(Me.properties.ExcelFilePath(), userRecord.GetIdNumber)
    Dim colTree As ExcelColumnNode = userRecord.GetCulumnNodeTree
    
    Try
      Me.excel.Open(filepath, True)
      
      ' 各月のレコードを読み込む
      userRecord.GetRecordDateTerm.MonthlyTerms.ForEach(
        Sub(m)
          Dim table As DataTable = userRecord.GetRecord(m.BeginDate.Month)
        
          If Me._cancel Then
            Return
          End If
          
          Dim sheetName As String = Me.properties.SheetName(m.BeginDate.Month)
          
          For day As Integer = m.BeginDate.Day To m.EndDate.Day
            ' 行を指定してデータを読み込む
            Dim row As Integer = day + Me.properties.FirstRow - 1
            Dim dict As IDictionary(Of String, String) = Me.excel.Read(row, filepath, sheetName, colTree)
            
            ' 読み込んだデータをDataTableの行にセットする
            Dim dataRow As DataRow = table.Rows(day - m.BeginDate.Day)
            dict.Keys.ForEach(
              Sub(k)
                If String.IsNullOrWhiteSpace(dict(k)) Then
                  Return
                End If
              
                'Log.out(userRecord.GetName & " " & m.BeginDate.Month.ToString & "/" & day.ToString & " " & k & ":" & dict(k)) 
                If table.Columns(k).DataType = GetType(Integer) Then
                  Dim v As Integer
                  If Integer.TryParse(dict(k), v) Then
                    dataRow(k) = v
                  End If
                ElseIf table.Columns(k).DataType = GetType(Double) Then
                  Dim v As Double
                  If Double.TryParse(dict(k), v) Then
                    dataRow(k) = v
                  End If
                Else
                  dataRow(k) = dict(k)
                End If
              End Sub)
          Next
        End Sub)
    Catch ex As Exception
      Throw ex
    Finally
      Me.excel.Close(filepath)
    End Try
  End Sub
  
'  ''' <summary>
'  ''' Excelプロパティの列設定から木構造の列コレクションを作成する。
'  ''' </summary>
'  Private Function CreateColumnNodeTree(properties As ExcelProperties) As ExcelColumnNode
'    Dim rootNode As New ExcelColumnNode(properties.WorkDayCol(), WORKDAY_COL_NAME, True)
'    
'    ' 各作業項目の列ノードを追加する
'    Dim idx As Integer = 1
'    While True
'      Dim param As ExcelProperties.WorkItemParams = properties.GetWorkItemParams(idx)
'      If param.Name = String.Empty Then
'        Exit While
'      End If
'      
'      Dim cntColNode As Nullable(Of ExcelColumnNode)
'      If param.WorkCountCol <> String.Empty Then
'        cntColNode = New ExcelColumnNode(param.WorkCountCol, param.Name & WORKCOUNT_COL_NAME)
'        rootNode.AddChild(cntColNode.Value)
'      End If
'      
'      If param.WorkTimeCol <> String.Empty Then
'        Dim timeColNode As New ExcelColumnNode(param.WorkTimeCol, param.Name & WORKTIME_COL_NAME)
'        If cntColNode.HasValue Then
'          cntColNode.Value.AddChild(timeColNode)
'        Else
'          rootNode.AddChild(timeColNode)          
'        End If
'      End If
'      
'      idx += 1
'    End While
'    
'    ' 備考の列ノードを追加する
'    rootNode.AddChild(New ExcelColumnNode(properties.NoteCol, properties.NoteName))
'    
'    Return rootNode
'  End Function
  
'  ''' <summary>
'  ''' 指定したExcelの列のコレクションからデータテーブルを定義し作成する。
'  ''' 列コレクションの各列に列名がついている必要がある。
'  ''' </summary>
'  ''' <param name="columnNodes"></param>
'  ''' <returns></returns>
'  Private Function CreateDataTable(columnNodes As List(Of ExcelColumnNode)) As DataTable
'    Dim table As New DataTable
'    AddColumnsToTable(columnNodes, table)
'    
'    Return table
'  End Function
'  
'  Private Sub AddColumnsToTable(nodes As List(Of ExcelColumnNode), table As DataTable)
'    nodes.ForEach(
'      Sub(n)
'        table.Columns.Add(CreateColumn(n.GetName))
'        AddColumnsToTable(n.GetChilds, table)
'      End Sub)
'  End Sub
'  
'  Public Function CreateColumn(name As String) As DataColumn
'    Dim col As New DataColumn
'    col.ColumnName = name
'    col.AutoIncrement = False
'		
'		Return col
'	End FUnction
End Class
