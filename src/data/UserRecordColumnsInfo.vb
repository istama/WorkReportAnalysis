'
' 日付: 2016/11/14
'

''' <summary>
''' レコードの列名を持つ構造体。
''' </summary>
Public Structure UserRecordColumnsInfo
  Public Const WORKDAY_COL_NAME   As String = "出勤日"
  
  ''' 各作業項目ごとの列情報を格納するリスト 
  Private ReadOnly workItems As List(Of WorkItemColumnsInfo)
  ''' 備考欄の列名 
  Public ReadOnly noteColName As String
  ''' 備考欄の列
  Public ReadOnly noteCol As String
  ''' 出勤日の列名 
  Public ReadOnly workDayColName As String
  ''' 出勤日の列
  Public ReadOnly workDayCol As String
  
  Public Sub New(properties As ExcelProperties)
    If properties Is Nothing Then Throw New NullReferenceException("properties is null")
    
    Me.workItems = New List(Of WorkItemColumnsInfo)
    
    Dim idx As Integer = 1
    While True 
      ' Excelのプロパティから新しい作業項目の設定が取得できたなら
      ' そこから列名を生成しリストに格納する。
      ' 取得できなかったらループを抜ける。
      Dim params As ExcelProperties.WorkItemParams = properties.GetWorkItemParams(idx)
      Dim colnames As WorkItemColumnsInfo? = WorkItemColumnsInfo.Create(params)
      If colnames.HasValue Then
        Me.workItems.Add(colnames.Value)
        idx += 1
      Else
        Exit While
      End If
    End While
    
    Me.noteColName = properties.NoteName
    Me.noteCol     = properties.NoteCol
    
    Me.workDayColName = WORKDAY_COL_NAME
    Me.workDayCol     = properties.WorkDayCol
  End Sub
  
  Public Function WorkItemList() As IList(Of WorkItemColumnsInfo)
    Return New Collections.ObjectModel.ReadOnlyCollection(Of WorkItemColumnsInfo)(Me.workItems)
  End Function
End Structure

''' <summary>
''' １つの作業項目あたりの列名。
''' </summary>
Public Structure WorkItemColumnsInfo
  Public Const WORKCOUNT_COL_NAME        As String = "件数"
  Public Const WORKTIME_COL_NAME         As String = "作業時間"
  Public Const WORKPRODUCTIVITY_COL_NAME As String = "生産性"
  
  Public ReadOnly WorkCountColName As String
  Public ReadOnly WorkCountCol As String
  Public ReadOnly WorkTimeColName As String
  Public ReadOnly WorkTimeCol As String
  Public ReadOnly WorkProductivityColName As String
  
  Private Sub New(params As ExcelProperties.WorkItemParams)
    If params.WorkCountCol <> String.Empty Then
      Me.WorkCountColName = params.Name & vbCrLf & WORKCOUNT_COL_NAME
      Me.WorkCountCol     = params.WorkCountCol
    Else
      Me.WorkCountColName = String.Empty
      Me.WorkCountCol     = String.Empty
    End If
    
    If params.WorkTimeCol <> String.Empty Then
      Me.WorkTimeColName = params.Name & vbCrLf & WORKTIME_COL_NAME
      Me.WorkTimeCol     = params.WorkTimeCol
    Else
      Me.WorkTimeColName = String.Empty
      Me.WorkTimeCol     = String.Empty
    End If
    
    If params.WorkCountCol <> String.Empty AndAlso params.WorkTimeCol <> String.Empty Then
      Me.WorkProductivityColName = params.Name & vbCrLf & WORKPRODUCTIVITY_COL_NAME
    Else
      Me.WorkProductivityColName = String.Empty
    End If
  End Sub
  
  Public Shared Function Create(params As ExcelProperties.WorkItemParams) As WorkItemColumnsInfo?
    If params.Name Is Nothing OrElse params.Name = String.Empty Then
      Return Nothing
    Else
      Return New WorkItemColumnsInfo(params)
    End If
  End Function
End Structure