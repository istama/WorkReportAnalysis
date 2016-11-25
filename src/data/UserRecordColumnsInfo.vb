'
' 日付: 2016/11/14
'
Imports System.Data
Imports System.Linq
Imports Common.COM
Imports Common.Extensions
Imports Common.IO

''' <summary>
''' レコードの列の情報。
''' </summary>
Public Structure ColumnInfo
  ''' 列名  
  Public ReadOnly name As String
  ''' 列（Excelでの列）
  Public ReadOnly col  As String
  ''' 値の型
  Public ReadOnly type As Type
  ''' この列をExcel読み込み時に結果に含めるかどうか
  Private ReadOnly containedToDataTable As Boolean
  
  Public Sub New(name As String, col As String, type As Type, Optional containedToDataTable As Boolean=True)
    If name Is Nothing Then Throw New ArgumentNullException("name is null")
    If col  Is Nothing Then Throw New ArgumentNullException("col is null")
    If type Is Nothing Then Throw New ArgumentNullException("type is null")
    
    Me.name = name
    Me.col  = col
    Me.type = type
    Me.containedToDataTable = containedToDataTable
  End Sub
  
  ''' <summary>
  ''' 列を作成する。
  ''' </summary>
  Public Function CreateDataColumn() As DataColumn
    Return ColumnInfo.CreateDataColumn(Me.name, Me.type)
  End Function
  
  ''' <summary>
  ''' 列を作成する。
  ''' </summary>
  Public Shared Function CreateDataColumn(name As String, type As Type) As DataColumn
    Dim col As New DataColumn
    col.ColumnName = name
    col.AutoIncrement = False
    col.DataType = type
		
		Return col
  End Function
  
  ''' <summary>
  ''' Excelを読み込むための列ノードを作成する。
  ''' </summary> 
  Public Function CreateExcelColumnNode() As ExcelColumnNode?
    If col <> String.Empty AndAlso name <> String.Empty Then
      Return New ExcelColumnNode(col, name, containedToDataTable)
    Else
      Return Nothing
    End If
  End Function
End Structure

''' <summary>
''' レコードのすべての列の情報を持つ構造体。
''' </summary>
Public Structure UserRecordColumnsInfo
  Public Const WORKDAY_COL_NAME As String = "出勤日"
  Public Const NAME_COL_NAME    As String = "名前"
  Public Const DATE_COL_NAME    As String = "日にち"
  
  Private ReadOnly properties As ExcelProperties
  
  ''' 備考欄の列情報
  Public ReadOnly noteColInfo As ColumnInfo
  ''' 出勤日の列情報 
  Public ReadOnly workDayColInfo As ColumnInfo
  
  Public Sub New(properties As ExcelProperties)
    If properties Is Nothing Then Throw New NullReferenceException("properties is null")
    
    Me.properties = properties
    
    Me.noteColInfo = New ColumnInfo(properties.NoteName, properties.NoteCol, GetType(String))
    Me.workDayColInfo = New ColumnInfo(WORKDAY_COL_NAME, properties.WorkDayCol, GetType(String), False)
  End Sub
  
  ''' <summary>
  ''' 作業項目のコレクションを返す。
  ''' </summary>
  Public Iterator Function WorkItemList() As IEnumerable(Of WorkItemColumnsInfo)
    For Each col As WorkItemColumnsInfo? In _
      Me.properties.GetWorkItemParamsEnumerable().
      Select(Function(params) WorkItemColumnsInfo.Create(params)).
      Where(Function(colInfo) colInfo.HasValue)
      Yield col.Value
    Next
  End Function
  
  ''' <summary>
  ''' テーブルを作成する。
  ''' </summary>
  Public Function CreateDataTable() As DataTable
    Return CreateDataTable(Nothing)
  End Function
  
  ''' <summary>
  ''' テーブルを作成する。
  ''' 一列目に引数の名前の列を追加することができる。
  ''' </summary>
  Public Function CreateDataTable(addedFirstColumnName As String) As DataTable
    Dim table As New DataTable
    
    ' この構造体の設定とは別に先頭の列に列を追加する
    If Not String.IsNullOrWhiteSpace(addedFirstColumnName) Then
      table.Columns.Add(ColumnInfo.CreateDataColumn(addedFirstColumnName, GetType(String)))
    End If
    
    ' 作業項目の列を追加する
    WorkItemList().ForEach(
      Sub(item)
        If Not String.IsNullOrWhiteSpace(item.WorkCountColInfo.Name) Then
          table.Columns.Add(item.WorkCountColInfo.CreateDataColumn())
        End If
        
        If Not String.IsNullOrWhiteSpace(item.WorkTimeColInfo.name) Then
          table.Columns.Add(item.WorkTimeColInfo.CreateDataColumn())
        End If
        
        If Not String.IsNullOrWhiteSpace(item.WorkProductivityColInfo.name) Then
          table.Columns.Add(item.WorkProductivityColInfo.CreateDataColumn)
        End If
      End Sub)
    
    ' 備考の列を追加する
    table.Columns.Add(Me.noteColInfo.CreateDataColumn)
    
    Return table
  End Function
  
  ''' <summary>
  ''' Excelを読み込むための列ノードのツリーを作成する。
  ''' </summary>
  Public Function CreateExcelColumnNodeTree() As ExcelColumnNode
    Dim rootNode As ExcelColumnNode? = Me.workDayColInfo.CreateExcelColumnNode()
    
    If Not rootNode.HasValue Then
      Throw New InvalidOperationException("excel.propertiesにWorkDayColの値が設定されていません。")  
    End If
    
    WorkItemList().ForEach(
      Sub(item)
        Dim node As ExcelColumnNode? = item.CreateExcelColumnNodeTree()
        If node.HasValue Then
          rootNode.Value.AddChild(node.Value)
        End If
      End Sub)
    
    Dim noteNode As ExcelColumnNode? = Me.noteColInfo.CreateExcelColumnNode()
    If noteNode.HasValue Then
      rootNode.Value.AddChild(noteNode.Value)
    End If
    
    Return rootNode.Value
  End Function
End Structure

''' <summary>
''' １つの作業項目あたりの列名。
''' </summary>
Public Structure WorkItemColumnsInfo
  Public Const WORKCOUNT_COL_NAME        As String = "件数"
  Public Const WORKTIME_COL_NAME         As String = "作業時間"
  Public Const WORKPRODUCTIVITY_COL_NAME As String = "生産性"
  
  ''' 作業件数の列情報
  Public ReadOnly WorkCountColInfo        As ColumnInfo
  ''' 作業時間の列情報
  Public ReadOnly WorkTimeColInfo         As ColumnInfo
  ''' 生産性の列情報
  Public ReadOnly WorkProductivityColInfo As ColumnInfo
  
  Private Sub New(params As ExcelProperties.WorkItemParams)
    ' 作業件数の列情報を作成
    If params.WorkCountCol <> String.Empty Then
      Me.WorkCountColInfo =
        New ColumnInfo(
          params.Name & WORKCOUNT_COL_NAME,
          params.WorkCountCol,
          GetType(Integer))
    Else
      Me.WorkCountColInfo =
        New ColumnInfo(
          String.Empty,
          String.Empty,
          GetType(Integer))
    End If
    
    ' 作業時間の列情報を作成
    If params.WorkTimeCol <> String.Empty Then
      Me.WorkTimeColInfo =
        New ColumnInfo(
          params.Name & WORKTIME_COL_NAME,
          params.WorkTimeCol,
          GetType(Double))
    Else
      Me.WorkTimeColInfo =
        New ColumnInfo(
          String.Empty,
          String.Empty,
          GetType(Double))
    End If
    
    ' 生産性の列情報を作成
    If params.WorkCountCol <> String.Empty AndAlso params.WorkTimeCol <> String.Empty Then
      Me.WorkProductivityColInfo =
        New ColumnInfo(
          params.Name & WORKPRODUCTIVITY_COL_NAME,
          String.Empty,
          GetType(Double))
    Else
      Me.WorkProductivityColInfo =
        New ColumnInfo(
          String.Empty,
          String.Empty,
          GetType(Double))
    End If
  End Sub
  
  Public Shared Function Create(params As ExcelProperties.WorkItemParams) As WorkItemColumnsInfo?
    If params.Name Is Nothing OrElse params.Name = String.Empty Then
      Return Nothing
    Else
      Return New WorkItemColumnsInfo(params)
    End If
  End Function
  
  ''' <summary>
  ''' Excelの列ノードのツリーを作成する。
  ''' </summary>
  Public Function CreateExcelColumnNodeTree() As ExcelColumnNode?
    Dim cntNode As ExcelColumnNode? = Me.WorkCountColInfo.CreateExcelColumnNode()
    Dim timeNode As ExcelColumnNode? = Me.WorkTimeColInfo.CreateExcelColumnNode()
    If cntNode.HasValue Then
      If timeNode.HasValue Then
        cntNode.Value.AddChild(timeNode.Value)
      End If
      
      Return cntNode
    ElseIf timeNode.HasValue
      Return timeNode
    Else
      Return NOthing  
    End If
  End Function
End Structure