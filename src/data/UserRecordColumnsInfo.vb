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
  ''' ダミー用の列  
  Private Shared ReadOnly _DUMMY As New ColumnInfo(String.Empty, String.Empty, GetType(String), False)
  
  ''' 列名  
  Public ReadOnly name As String
  ''' 列（Excelでの列）
  Public ReadOnly col  As String
  ''' 値の型
  Public ReadOnly type As Type
  ''' この列をExcel読み込み時に結果に含めるかどうか
  Private ReadOnly containedToDataTable As Boolean
  
  ''' <summary>
  ''' ファクトリーメソッド。
  ''' </summary>
  Public Shared Function Create(name As String, col As String, type As Type, Optional containedToDataTable As Boolean=True) As ColumnInfo
    If name Is Nothing Then Throw New ArgumentNullException("name is null")
    If col  Is Nothing Then Throw New ArgumentNullException("col is null")
    If type Is Nothing Then Throw New ArgumentNullException("type is null")
    
    ' 列の名前が空なので例外を投げる
    If String.IsNullOrWhiteSpace(name) Then Throw New ArgumentException("name is empty")
    ' 列の値が不正なので例外を投げる
    If Not Cell.ValidColumn(col)       Then Throw New ArgumentException("col is invalid / " & col)    
    
    Return New ColumnInfo(name, col, type, containedToDataTable)
  End Function
  
  ''' <summary>
  ''' ファクトリーメソッド。
  ''' Excelを読み込まない列を作成する場合に使用する。
  ''' </summary>
  Public Shared Function Create(name As String, type As Type, Optional containedToDataTable As Boolean=True) As ColumnInfo
    If name Is Nothing Then Throw New ArgumentNullException("name is null")
    If type Is Nothing Then Throw New ArgumentNullException("type is null")
    
    ' 列の名前が空なので例外を投げる
    If String.IsNullOrWhiteSpace(name) Then Throw New ArgumentException("name is empty")
    
    Return New ColumnInfo(name, String.Empty, type, containedToDataTable)    
  End Function
  
  ''' <summary>
  ''' ファクトリーメソッド。
  ''' ダミー用の列情報を作成する。
  ''' </summary>
  Public Shared Function Dummy() As ColumnInfo
    Return _DUMMY  
  End Function
  
  Private Sub New(name As String, col As String, type As Type, Optional containedToDataTable As Boolean=True)
    Me.name = name
    Me.col  = col
    Me.type = type
    Me.containedToDataTable = containedToDataTable
  End Sub
  
  ''' <summary>
  ''' 列を作成する。
  ''' ダミー用の列の場合はnullを返す。
  ''' </summary>
  Public Function CreateDataColumn() As DataColumn
    Return ColumnInfo.CreateDataColumn(Me.name, Me.type)
  End Function
  
  ''' <summary>
  ''' 列を作成する。
  ''' </summary>
  Public Shared Function CreateDataColumn(name As String, type As Type) As DataColumn
    If String.IsNullOrWhiteSpace(name) OrElse type Is Nothing Then
      Return Nothing
    Else
      Dim col As New DataColumn
      col.ColumnName    = name
      col.AutoIncrement = False
      col.DataType      = type
  		
  		Return col     
    End If
  End Function
  
  ''' <summary>
  ''' Excelを読み込むための列ノードを作成する。
  ''' 列情報が完全でない場合はnullを返す。
  ''' </summary> 
  Public Function CreateExcelColumnNode() As Nullable(Of ExcelColumnNode)
    If col = String.Empty Then
      Return Nothing
    Else
      Return New ExcelColumnNode(col, name, containedToDataTable)  
    End If
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
  
  ''' <summary>
  ''' ファクトリーメソッド。
  ''' 作業項目名が空の場合、もしくはExcelの列が１つも設定されていない場合は例外を投げる。
  ''' </summary>
  Public Shared Function Create(params As ExcelProperties.WorkItemParams) As WorkItemColumnsInfo
    If String.IsNullOrWhiteSpace(params.Name)        Then Throw New ArgumentException("Excelプロパティの作業項目名が設定されていません。")
    If String.IsNullOrWhiteSpace(params.WorkCountCol) AndAlso _
       String.IsNullOrWhiteSpace(params.WorkTimeCol) Then Throw New ArgumentException("Excelプロパティの列情報が設定されていません。")
      
    Return New WorkItemColumnsInfo(params)
  End Function
  
  Private Sub New(params As ExcelProperties.WorkItemParams)
    ' 作業件数の列情報を作成
    If params.WorkCountCol <> String.Empty Then
      Me.WorkCountColInfo =
        ColumnInfo.Create(
          params.Name & WORKCOUNT_COL_NAME,
          params.WorkCountCol,
          GetType(Integer))
    Else
      Me.WorkCountColInfo = ColumnInfo.Dummy
    End If
    
    ' 作業時間の列情報を作成
    If params.WorkTimeCol <> String.Empty Then
      Me.WorkTimeColInfo =
        ColumnInfo.Create(
          params.Name & WORKTIME_COL_NAME,
          params.WorkTimeCol,
          GetType(Double))
    Else
      Me.WorkTimeColInfo = ColumnInfo.Dummy
    End If
    
    ' 生産性の列情報を作成
    If params.WorkCountCol <> String.Empty AndAlso params.WorkTimeCol <> String.Empty Then
      Me.WorkProductivityColInfo =
        ColumnInfo.Create(
          params.Name & WORKPRODUCTIVITY_COL_NAME,
          GetType(Double))
    Else
      Me.WorkProductivityColInfo = ColumnInfo.Dummy
    End If
  End Sub
  
  ''' <summary>
  ''' Excelの列ノードのツリーを作成する。
  ''' </summary>
  Public Function CreateExcelColumnNodeTree() As ExcelColumnNode
    ' 作業件数と作業時間の列情報を生成
    Dim cntNode  As Nullable(Of ExcelColumnNode) = Me.WorkCountColInfo.CreateExcelColumnNode()
    Dim timeNode As Nullable(Of ExcelColumnNode) = Me.WorkTimeColInfo.CreateExcelColumnNode()
    
    ' 列情報が有効な場合、親子関係に接続して返す。
    If cntNode.HasValue Then
      If timeNode.HasValue Then
        cntNode.Value.AddChild(timeNode.Value)
      End If
      
      Return cntNode.Value
    ElseIf timeNode.HasValue
      Return timeNode.Value
    Else
      ' 通常起こらないエラー
      Throw New InvalidOperationException("オブジェクトの状態が不正です。作業項目の列情報が１つも設定されていません。")
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
  
  ''' 作業項目の列情報のリスト
  Private workItemList As List(Of WorkItemColumnsInfo)
  
  ''' <summary>
  ''' ファクトリーメソッド。
  ''' </summary>
  Public Shared Function Create(properties As ExcelProperties) As UserRecordColumnsInfo
    If properties Is Nothing Then Throw New ArgumentNullException("properties is null")
    
    If String.IsNullOrWhiteSpace(properties.NoteName) Then
      Throw New ArgumentException("Excelプロパティファイルの NoteName が設定されていません。")
    ElseIf String.IsNullOrWhiteSpace(properties.NoteCol)
      Throw New ArgumentException(
        "Excelプロパティファイルの NoteCol が設定されていません。" & vbCrLf &
        "NoteCol とは備考欄の列のことです。")
    ElseIf String.IsNullOrWhiteSpace(properties.WorkDayCol)
      Throw New ArgumentException(
        "Excelプロパティファイルの WorkDayCol が設定されていません。" & vbCrLf &
        "WorkDayCol とは出勤日の列のことです。")
    End If
    
    Return New UserRecordColumnsInfo(properties)
  End Function
  
  Private Sub New(properties As ExcelProperties)
    If properties Is Nothing Then Throw New NullReferenceException("properties is null")
    
    Me.properties = properties
    
    Me.noteColInfo    = ColumnInfo.Create(properties.NoteName, properties.NoteCol,    GetType(String))
    Me.workDayColInfo = ColumnInfo.Create(WORKDAY_COL_NAME,    properties.WorkDayCol, GetType(String), False)
  End Sub
  
  ''' <summary>
  ''' 作業項目のコレクションを返す。
  ''' </summary>
  Public Function WorkItems() As IEnumerable(Of WorkItemColumnsInfo)
    ' 列情報のリストがまだ作成されていない場合
    If Me.workItemList Is Nothing Then
      Dim items As IEnumerable(Of WorkItemColumnsInfo) =
        Me.properties.GetWorkItemParamsEnumerable().
        Select(Function(params) WorkItemColumnsInfo.Create(params))
      
      Me.workItemList = New List(Of WorkItemColumnsInfo)(items)
    End If
    
    Return Me.workItemList
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
    For Each item As WorkItemColumnsInfo In WorkItems()
      If Not String.IsNullOrWhiteSpace(item.WorkCountColInfo.Name) Then
        table.Columns.Add(item.WorkCountColInfo.CreateDataColumn())
      End If
      
      If Not String.IsNullOrWhiteSpace(item.WorkTimeColInfo.name) Then
        table.Columns.Add(item.WorkTimeColInfo.CreateDataColumn())
      End If
      
      If Not String.IsNullOrWhiteSpace(item.WorkProductivityColInfo.name) Then
        table.Columns.Add(item.WorkProductivityColInfo.CreateDataColumn)
      End If      
    Next
    
    ' 備考の列を追加する
    table.Columns.Add(Me.noteColInfo.CreateDataColumn)
    
    Return table
  End Function
  
  ''' <summary>
  ''' Excelを読み込むための列ノードのツリーを作成する。
  ''' </summary>
  Public Function CreateExcelColumnNodeTree() As ExcelColumnNode
    ' 出勤日の列を作成
    Dim rootNode As Nullable(Of ExcelColumnNode) = Me.workDayColInfo.CreateExcelColumnNode()
    
    If Not rootNode.HasValue Then
      Throw New InvalidOperationException(
        "excel.propertiesファイル の WorkDayCol に 出勤日の列　の値を設定してください。" & vbCrLf &
        "これを設定することでExcelファイルの読み込みが速くなります。")
    End If
    
    ' 各作業項目の列を１つずつ返し、出勤日の列の子要素とする
    For Each item As WorkItemColumnsInfo In WorkItems()
      Dim node As Nullable(Of ExcelColumnNode) = item.CreateExcelColumnNodeTree()
      
      If node.HasValue Then
        rootNode.Value.AddChild(node.Value)
      End If      
    Next
    
    ' 備考列を生成し、出勤日の列の子要素とする
    Dim noteNode As Nullable(Of ExcelColumnNode) = Me.noteColInfo.CreateExcelColumnNode()
    If noteNode.HasValue Then
      rootNode.Value.AddChild(noteNode.Value)
    End If
    
    Return rootNode.Value
  End Function
End Structure

